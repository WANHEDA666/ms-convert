using ms_converter.service.errors;
using NetOffice.WordApi;
using NetOffice.PowerPointApi;

namespace ms_converter.service;

public sealed class Converter(Storage storage)
{
    public void ConvertToPdf(string srcPath)
    {
        var dstPath = storage.GetResultPdfPath();
        Exception? err = null;
        var thread = new Thread(() =>
        {
            try
            {
                var ext = Path.GetExtension(srcPath).TrimStart('.').ToLowerInvariant();
                switch (ext)
                {
                    case "doc":
                    case "docx":
                    case "odt":
                    case "txt":
                        ConvertWord(srcPath, dstPath);
                        break;
                    case "ppt":
                    case "pptx":
                    case "pps":
                    case "ppsx":
                    case "pptm":
                    case "pot":
                    case "odp":
                        ConvertPowerPoint(srcPath, dstPath);
                        break;
                    default:
                        throw new NotSupportedException($"Unsupported extension: {ext}");
                }
            }
            catch (System.Runtime.InteropServices.COMException com)
            {
                throw new OfficeApiException(com.Message);
            }
            catch (Exception ex) when (ex.Source?.Contains("Word", StringComparison.OrdinalIgnoreCase) == true 
                                       || ex.Source?.Contains("PowerPoint", StringComparison.OrdinalIgnoreCase) == true
                                       || (ex.StackTrace?.Contains("NetOffice.WordApi") ?? false)
                                       || (ex.StackTrace?.Contains("NetOffice.PowerPointApi") ?? false))
            {
                throw new OfficeApiException(ex.Message);
            }
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();
        if (err != null) throw err;
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
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
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
            pres = ppt.Presentations.Open(srcPath, NetOffice.OfficeApi.Enums.MsoTriState.msoFalse, NetOffice.OfficeApi.Enums.MsoTriState.msoFalse, 
                NetOffice.OfficeApi.Enums.MsoTriState.msoFalse);
            pres.SaveAs(dstPath, NetOffice.PowerPointApi.Enums.PpSaveAsFileType.ppSaveAsPDF);
            pres.Close();
            ppt.Quit();
        }
        finally
        {
            try { pres?.Dispose(); } catch { }
            try { ppt?.Dispose(); } catch { }
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}