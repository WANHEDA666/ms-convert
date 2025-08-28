using NetOffice.WordApi;
using NetOffice.PowerPointApi;

namespace ms_converter.service;

public sealed class Converter
{
    public void ConvertToPdf(string srcPath)
    {
        var dstPath = Path.Combine(Path.Combine(AppContext.BaseDirectory, "result"), "file.pdf");
        var ext = Path.GetExtension(srcPath).TrimStart('.').ToLowerInvariant();
        switch (ext)
        {
            case "doc":
            case "docx":
                ConvertWord(srcPath, dstPath);
                break;
            case "ppt":
            case "pptx":
                ConvertPowerPoint(srcPath, dstPath);
                break;
        }
    }

    private void ConvertWord(string srcPath, string dstPath)
    {
        NetOffice.WordApi.Application? word = null;
        Document? doc = null;
        try
        {
            word = new NetOffice.WordApi.Application { Visible = false };
            word.DisplayAlerts = NetOffice.WordApi.Enums.WdAlertLevel.wdAlertsNone;
            doc = word.Documents.Open(srcPath);
            doc.SaveAs2(dstPath, NetOffice.WordApi.Enums.WdSaveFormat.wdFormatPDF);
            doc.Close(NetOffice.WordApi.Enums.WdSaveOptions.wdDoNotSaveChanges);
            word.Quit();
        }
        finally
        {
            try { doc?.Dispose(); } catch { }
            try { word?.Dispose(); } catch { }
        }
    }

    private void ConvertPowerPoint(string srcPath, string dstPath)
    {
        NetOffice.PowerPointApi.Application? ppt = null;
        Presentation? pres = null;
        try
        {
            ppt = new NetOffice.PowerPointApi.Application();
            ppt.DisplayAlerts = NetOffice.PowerPointApi.Enums.PpAlertLevel.ppAlertsNone;
            pres = ppt.Presentations.Open(
                srcPath,
                NetOffice.OfficeApi.Enums.MsoTriState.msoFalse,
                NetOffice.OfficeApi.Enums.MsoTriState.msoFalse,
                NetOffice.OfficeApi.Enums.MsoTriState.msoFalse
            );
            pres.SaveAs(
                dstPath,
                NetOffice.PowerPointApi.Enums.PpSaveAsFileType.ppSaveAsPDF,
                NetOffice.OfficeApi.Enums.MsoTriState.msoFalse
            );
            pres.Close();
            ppt.Quit();
        }
        finally
        {
            try { pres?.Dispose(); } catch { }
            try { ppt?.Dispose(); } catch { }
        }
    }
}