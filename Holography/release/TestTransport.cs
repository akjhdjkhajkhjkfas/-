using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using HolographyPreview;

class TestTransport
{
    static void Check(bool condition,string message) { if(!condition) throw new Exception(message); }
    static async Task<byte[]> Receive(NetworkStream stream)
    {
        var data=new byte[320]; int offset=0;
        while(offset<320) { int count=await stream.ReadAsync(data,offset,320-offset); if(count==0) throw new Exception("short packet"); offset+=count; }
        return data;
    }
    static TcpListener Listen() { var server=new TcpListener(IPAddress.Loopback,0); server.Start(); return server; }
    static int Port(TcpListener server) { return ((IPEndPoint)server.LocalEndpoint).Port; }
    static void Packing()
    {
        var matrix=new bool[40,64];
        Check(PacketEncoder.Encode(matrix,BitOrder.MostSignificantFirst).All(b=>b==0),"all-zero packet");
        for(int r=0;r<40;r++) for(int c=0;c<64;c++)
        {
            matrix[r,c]=true;
            var packet=PacketEncoder.Encode(matrix,BitOrder.MostSignificantFirst);
            Check(packet.Length==320,"length");
            for(int i=0;i<320;i++) Check(packet[i]==(i==r*8+c/8 ? 1<<(7-c%8) : 0),"one-hot mapping");
            matrix[r,c]=false;
        }
        for(int r=0;r<40;r++) for(int c=0;c<64;c++) matrix[r,c]=true;
        Check(PacketEncoder.Encode(matrix,BitOrder.MostSignificantFirst).All(b=>b==255),"all-one packet");
        matrix=new bool[40,64]; matrix[0,0]=true; matrix[0,7]=true; matrix[1,0]=true; matrix[39,63]=true;
        var known=PacketEncoder.Encode(matrix,BitOrder.MostSignificantFirst);
        Check(known[0]==0x81 && known[8]==0x80 && known[319]==0x01,"known vector");
        Console.WriteLine("PASS: all 2560 one-hot channels, MSB-first row order, 320-byte length, zero/one and known vectors");
    }
    static async Task Repeated()
    {
        var server=Listen();
        try
        {
            var expected=PacketEncoder.Encode(Codebook.Demo(),BitOrder.MostSignificantFirst);
            var peer=Task.Run(async delegate {
                using(var client=await server.AcceptTcpClientAsync())
                {
                    var stream=client.GetStream();
                    for(int n=0;n<2;n++)
                    {
                        Check((await Receive(stream)).SequenceEqual(expected),"wire payload differs from preview");
                        byte[] part=Encoding.ASCII.GetBytes(n==0?"O":"ER");
                        await stream.WriteAsync(part,0,part.Length); await Task.Delay(40);
                        part=Encoding.ASCII.GetBytes(n==0?"K":"ROR");
                        await stream.WriteAsync(part,0,part.Length);
                    }
                }
            });
            using(var session=new TcpSession())
            {
                await session.ConnectAsync("127.0.0.1",Port(server),1000);
                Check(await session.SendAsync(expected,true,1000)=="OK","split OK");
                Check(await session.SendAsync(expected,true,1000)=="ERROR","split ERROR");
            }
            await peer;
            Console.WriteLine("PASS: persistent TCP connection, repeated exact 320-byte sends, fragmented OK/ERROR replies");
        }
        finally { server.Stop(); }
    }
    static async Task TimeoutAndDuplicate()
    {
        var server=Listen();
        try
        {
            var peer=Task.Run(async delegate {
                using(var client=await server.AcceptTcpClientAsync())
                {
                    var stream=client.GetStream(); await Receive(stream);
                    int extra=await stream.ReadAsync(new byte[1],0,1);
                    Check(extra==0,"unexpected automatic retry");
                }
            });
            using(var session=new TcpSession())
            {
                await session.ConnectAsync("127.0.0.1",Port(server),1000);
                var pending=session.SendAsync(new byte[320],true,180);
                try { await session.SendAsync(new byte[320],true,1000); throw new Exception("duplicate allowed"); }
                catch(InvalidOperationException) { }
                try { await pending; throw new Exception("timeout absent"); } catch(TimeoutException) { }
                Check(!session.Connected,"timed-out session remains open");
            }
            await peer;
            Console.WriteLine("PASS: reply timeout closes connection, concurrent sends rejected, no automatic retry");
        }
        finally { server.Stop(); }
    }
    static async Task NoReplyAndDisconnect()
    {
        foreach(bool noReply in new[]{true,false})
        {
            var server=Listen();
            try
            {
                var peer=Task.Run(async delegate {
                    using(var client=await server.AcceptTcpClientAsync())
                        Check((await Receive(client.GetStream())).Length==320,"receive");
                });
                using(var session=new TcpSession())
                {
                    await session.ConnectAsync("127.0.0.1",Port(server),1000);
                    if(noReply) Check(await session.SendAsync(new byte[320],false,1000)=="SENT","no-reply mode");
                    else
                    {
                        try { await session.SendAsync(new byte[320],true,1000); throw new Exception("disconnect reported success"); }
                        catch(IOException) { }
                        Check(!session.Connected,"disconnected session stays open");
                    }
                }
                await peer;
            }
            finally { server.Stop(); }
        }
        Console.WriteLine("PASS: optional no-reply mode; remote close cannot be reported as execution success");
    }
    static async Task InvalidReplyAndReconnect()
    {
        var server=Listen();
        try
        {
            var peer=Task.Run(async delegate {
                using(var client=await server.AcceptTcpClientAsync())
                {
                    var stream=client.GetStream(); await Receive(stream);
                    var invalid=Encoding.ASCII.GetBytes("UNKNOWN"); await stream.WriteAsync(invalid,0,invalid.Length);
                }
                using(var client=await server.AcceptTcpClientAsync())
                {
                    var stream=client.GetStream(); await Receive(stream);
                    var ok=Encoding.ASCII.GetBytes("OK"); await stream.WriteAsync(ok,0,ok.Length);
                }
            });
            using(var session=new TcpSession())
            {
                await session.ConnectAsync("127.0.0.1",Port(server),1000);
                try { await session.SendAsync(new byte[320],true,1000); throw new Exception("invalid reply accepted"); } catch(IOException) { }
                Check(!session.Connected,"invalid reply connection stays open");
                await session.ConnectAsync("127.0.0.1",Port(server),1000);
                Check(await session.SendAsync(new byte[320],true,1000)=="OK","reconnect");
            }
            await peer;
            Console.WriteLine("PASS: unknown reply rejected, reconnect restores sending");
        }
        finally { server.Stop(); }
        using(var session=new TcpSession())
        {
            try { await session.SendAsync(new byte[320],true,1000); throw new Exception("disconnected send accepted"); } catch(IOException) { }
            try { await session.ConnectAsync("invalid",5001,1000); throw new Exception("invalid IP accepted"); } catch(ArgumentException) { }
        }
        Console.WriteLine("PASS: disconnected send and invalid IP rejected");
    }
    static async Task Run() { Packing(); await Repeated(); await TimeoutAndDuplicate(); await NoReplyAndDisconnect(); await InvalidReplyAndReconnect(); }
    static int Main() { try { Run().GetAwaiter().GetResult(); return 0; } catch(Exception ex) { Console.WriteLine("FAIL: "+ex); return 1; } }
}
