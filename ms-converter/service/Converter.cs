using System.Runtime.InteropServices;
using ms_converter.service.errors;
using NetOffice.OfficeApi.Enums;
using NetOffice.WordApi;
using NetOffice.PowerPointApi;
using NetOffice.PowerPointApi.Enums;
using NetOffice.WordApi.Enums;

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
            catch (COMException com)
            {
                err = new LocalOfficeApiException(com.Message);
            }
            catch (Exception ex) when (ex.Source?.Contains("Word", StringComparison.OrdinalIgnoreCase) == true 
                                       || ex.Source?.Contains("PowerPoint", StringComparison.OrdinalIgnoreCase) == true
                                       || (ex.StackTrace?.Contains("NetOffice.WordApi") ?? false)
                                       || (ex.StackTrace?.Contains("NetOffice.PowerPointApi") ?? false))
            {
                err = new LocalOfficeApiException(ex.Message);
            }
            catch (Exception ex)
            {
                err = ex;
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
            word.DisplayAlerts = WdAlertLevel.wdAlertsNone;
            doc = word.Documents.OpenNoRepairDialog(
                fileName: srcPath,
                confirmConversions: MsoTriState.msoFalse,
                readOnly: MsoTriState.msoTrue,
                addToRecentFiles: MsoTriState.msoFalse,
                passwordDocument: "FAKE",
                passwordTemplate: "FAKE",
                revert: MsoTriState.msoCTrue,
                writePasswordDocument: "FAKE",
                writePasswordTemplate: "FAKE",
                format: WdOpenFormat.wdOpenFormatAuto,
                encoding: MsoEncoding.msoEncodingAutoDetect,
                visible: MsoTriState.msoFalse,
                openAndRepair: MsoTriState.msoTrue,
                documentDirection: WdDocumentDirection.wdLeftToRight,
                noEncodingDialog: MsoTriState.msoTrue
            );

            doc.SaveAs2(dstPath, WdSaveFormat.wdFormatPDF);
            doc.Close(WdSaveOptions.wdDoNotSaveChanges);
            doc.Dispose();
            doc = null;
        }
        catch (COMException com)
        {
            throw new LocalOfficeApiException(com.Message);
        }
        catch (Exception ex) when ((ex.Source?.Contains("Word", StringComparison.OrdinalIgnoreCase) ?? false) || (ex.StackTrace?.Contains("NetOffice.WordApi") ?? false))
        {
            throw new LocalOfficeApiException(ex.Message);
        }
        finally
        {
            try { doc?.Close(WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { doc?.Dispose(); } catch { }
            try { word?.Quit(); } catch { }
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
            ppt.DisplayAlerts = PpAlertLevel.ppAlertsNone;
            var openPath = srcPath + "::FAKE::FAKE";
            try
            {
                pres = ppt.Presentations.Open(
                    fileName: openPath,
                    readOnly: MsoTriState.msoTrue,
                    untitled: MsoTriState.msoFalse,
                    withWindow: MsoTriState.msoFalse
                );
            }
            catch
            {
                pres = ppt.Presentations.Open2007(
                    fileName: openPath,
                    readOnly: MsoTriState.msoTrue,
                    untitled: MsoTriState.msoFalse,
                    withWindow: MsoTriState.msoFalse,
                    openAndRepair: MsoTriState.msoTrue
                );
            }

            pres.SaveAs(dstPath, PpSaveAsFileType.ppSaveAsPDF);
            pres.Close();
            pres.Dispose();
            pres = null;
        }
        catch (COMException com)
        {
            throw new LocalOfficeApiException(com.Message);
        }
        catch (Exception ex) when ((ex.Source?.Contains("PowerPoint", StringComparison.OrdinalIgnoreCase) ?? false) || (ex.StackTrace?.Contains("NetOffice.PowerPointApi") ?? false))
        {
            throw new LocalOfficeApiException(ex.Message);
        }
        finally
        {
            try { pres?.Close(); } catch { }
            try { pres?.Dispose(); } catch { }
            try { ppt?.Quit(); } catch { }
            try { ppt?.Dispose(); } catch { }
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}