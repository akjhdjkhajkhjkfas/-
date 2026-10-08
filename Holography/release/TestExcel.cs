using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using HolographyPreview;

class TestExcel
{
    [STAThread]
    static int Main(string[] args)
    {
        object app=null, books=null, book=null, sheets=null, sheet=null, range=null;
        string output=Path.GetFullPath(args[0]);
        try
        {
            app=Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            dynamic excel=app;
            excel.Visible=false; excel.DisplayAlerts=false; excel.AutomationSecurity=3;
            books=excel.Workbooks; book=((dynamic)books).Add();
            sheets=((dynamic)book).Worksheets; sheet=((dynamic)sheets).Item[1];
            range=((dynamic)sheet).Range["A1", "BL40"];
            var data=Codebook.Demo(); var values=new object[40,64];
            for(int r=0;r<40;r++) for(int c=0;c<64;c++) values[r,c]=data[r,c]?1:0;
            ((dynamic)range).Value2=values;
            ((dynamic)book).SaveAs(Path.Combine(output,"test_40x64.xlsx"),51);
            ((dynamic)book).SaveAs(Path.Combine(output,"test_40x64.xls"),56);
            ((dynamic)book).Close(false); Marshal.FinalReleaseComObject(book); book=null;
            excel.Quit();
            foreach(string extension in new[]{".xls",".xlsx"})
            {
                var result=Codebook.Read(Path.Combine(output,"test_40x64"+extension));
                if(!data.Cast<bool>().SequenceEqual(result.Cast<bool>())) throw new Exception("Excel mismatch "+extension);
                Console.WriteLine("PASS: "+extension+" import, all 2560 values match");
            }
            try { Codebook.Read(Path.GetFullPath(args[1])); throw new Exception("Legacy dimensions accepted"); }
            catch(FormatException ex) { Console.WriteLine("PASS: original 72x32 XLS rejected explicitly: "+ex.Message); }
            return 0;
        }
        catch(Exception ex) { Console.Error.WriteLine(ex); return 1; }
        finally
        {
            if(book!=null) try { ((dynamic)book).Close(false); } catch {}
            if(app!=null) try { ((dynamic)app).Quit(); } catch {}
            foreach(var obj in new[]{range,sheet,sheets,book,books,app})
                if(obj!=null && Marshal.IsComObject(obj)) Marshal.FinalReleaseComObject(obj);
        }
    }
}
