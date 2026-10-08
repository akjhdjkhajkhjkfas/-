using System;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace HolographyPreview
{
    public enum BitOrder { MostSignificantFirst }

    public static class PacketEncoder
    {
        // Board's row-major, MSB-first layout confirmed by the user.
        public static byte[] Encode(bool[,] data, BitOrder order)
        {
            if (data.GetLength(0) != Codebook.Rows || data.GetLength(1) != Codebook.Columns)
                throw new ArgumentException("需要 40 行 × 64 列数据");
            var bytes = new byte[320];
            for (int r = 0; r < 40; r++)
                for (int c = 0; c < 64; c++)
                    if (data[r,c]) bytes[r*8+c/8] |= (byte)(1 << (7-c%8));
            return bytes;
        }
    }

    public sealed class TcpSession : IDisposable
    {
        TcpClient client;
        readonly SemaphoreSlim sendGate = new SemaphoreSlim(1,1);
        public bool Connected { get { return client != null && client.Connected; } }

        public async Task ConnectAsync(string address, int port, int timeoutMs)
        {
            IPAddress ip;
            if (!IPAddress.TryParse(address, out ip)) throw new ArgumentException("设备 IP 地址无效");
            if (port < 1 || port > 65535) throw new ArgumentException("端口需要在 1–65535 之间");
            Close();
            var next = new TcpClient(ip.AddressFamily);
            client = next;
            try
            {
                var connect = next.ConnectAsync(ip, port);
                if (await Task.WhenAny(connect, Task.Delay(timeoutMs)) != connect)
                {
                    next.Close();
                    try { await connect; } catch { }
                    throw new TimeoutException("连接超时");
                }
                await connect;
                if (client != next) throw new IOException("连接已取消");
                next.NoDelay = true;
            }
            catch { if (client == next) Close(); else next.Close(); throw; }
        }

        public async Task<string> SendAsync(byte[] packet, bool requireReply, int timeoutMs)
        {
            if (packet == null || packet.Length == 0) throw new ArgumentException("发送数据为空");
            if (!await sendGate.WaitAsync(0)) throw new InvalidOperationException("上一条发送尚未完成");
            TcpClient current = client;
            try
            {
                if (current == null || !current.Connected) throw new IOException("设备尚未连接");
                var stream = current.GetStream();
                var write = stream.WriteAsync(packet,0,packet.Length);
                if (await Task.WhenAny(write,Task.Delay(timeoutMs)) != write)
                {
                    current.Close();
                    try { await write; } catch { }
                    throw new TimeoutException("发送超时，连接已关闭；不会自动重发");
                }
                await write;
                if (!requireReply) return "SENT";
                var read = ReadReplyAsync(stream);
                if (await Task.WhenAny(read,Task.Delay(timeoutMs)) != read)
                {
                    current.Close();
                    try { await read; } catch { }
                    throw new TimeoutException("等待设备回执超时；数据可能已执行，连接已关闭，不会自动重发");
                }
                return await read;
            }
            catch { if (client == current) Close(); throw; }
            finally { sendGate.Release(); }
        }

        static async Task<string> ReadReplyAsync(NetworkStream stream)
        {
            var text = new StringBuilder();
            var buffer = new byte[32];
            while (text.Length < 64)
            {
                int n = await stream.ReadAsync(buffer,0,buffer.Length);
                if (n == 0) throw new IOException("设备在回执前断开连接，无法确认执行结果");
                text.Append(Encoding.ASCII.GetString(buffer,0,n));
                string reply = text.ToString().Trim();
                if (reply == "OK" || reply == "ERROR") return reply;
                if (!"OK".StartsWith(reply,StringComparison.Ordinal) && !"ERROR".StartsWith(reply,StringComparison.Ordinal))
                    throw new IOException("无法识别设备回执："+reply);
            }
            throw new IOException("设备回执过长");
        }

        public void Close() { var current = client; client = null; if (current != null) current.Close(); }
        public void Dispose() { Close(); }
    }
}
