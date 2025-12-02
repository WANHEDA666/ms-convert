using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using ms_converter.service.errors;
using NetOffice.OfficeApi.Enums;
using NetOffice.WordApi;
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

            KillAllComSurrogates();
        }
    }

    private static void KillAllComSurrogates()
    {
        foreach (var p in Process.GetProcessesByName("dllhost"))
        {
            try { p.Kill(entireProcessTree: true); }
            catch { /* ignore */ }
        }
    }

    private void ConvertPowerPoint(string srcPath, string dstPath)
    {
        NetOffice.PowerPointApi.Application? ppt = null;
        NetOffice.PowerPointApi.Presentation? pres = null;
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
            try
            {
                if (pres != null)
                {
                    try { pres.Close(); } catch { }
                    pres.Dispose();
                }
            }
            catch { }

            try
            {
                if (ppt != null)
                {
                    try { ppt.Quit(); } catch { }
                    ppt.Dispose();
                }
            }
            catch { }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    
    public void ConvertToHtml(string srcPath)
    {
        var dstHtmlPath = storage.GetResultHtmlPath();
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
                        ConvertWordToHtml(srcPath, dstHtmlPath);
                        break;
                    case "ppt":
                    case "pptx":
                    case "pps":
                    case "ppsx":
                    case "pptm":
                    case "pot":
                    case "odp":
                        ConvertPowerPointToHtml(srcPath, dstHtmlPath);
                        break;

                    default:
                        throw new NotSupportedException($"Unsupported extension for HTML: {ext}");
                }
            }
            catch (COMException com)
            {
                err = new LocalOfficeApiException(com.Message);
            }
            catch (Exception ex) when ((ex.Source?.Contains("Word", StringComparison.OrdinalIgnoreCase) ?? false) || (ex.StackTrace?.Contains("NetOffice.WordApi") ?? false))
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

    private void ConvertWordToHtml(string srcPath, string dstHtmlPath)
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

            doc.WebOptions.Encoding = MsoEncoding.msoEncodingUTF8;
            doc.SaveAs2(dstHtmlPath, WdSaveFormat.wdFormatFilteredHTML);
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

    private void ConvertPowerPointToHtml(string srcPath, string dstHtmlPath)
    {
        NetOffice.PowerPointApi.Application? ppt = null;
        NetOffice.PowerPointApi.Presentation? pres = null;
        NetOffice.PowerPointApi.Slides? slidesCollection = null;

        try
        {
            var dstDir = Path.GetDirectoryName(dstHtmlPath) ?? ".";
            Directory.CreateDirectory(dstDir);

            var imagesDir = Path.Combine(dstDir, Path.GetFileNameWithoutExtension(dstHtmlPath) + ".files");
            Directory.CreateDirectory(imagesDir);

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

            int slideWidth = (int)pres.PageSetup.SlideWidth;
            int slideHeight = (int)pres.PageSetup.SlideHeight;

            var slidesResult = new List<(string ImgPath, string Text)>();

            slidesCollection = pres.Slides;
            int slidesCount = slidesCollection.Count;

            for (int i = 1; i <= slidesCount; i++)
            {
                NetOffice.PowerPointApi.Slide? slide = null;
                NetOffice.PowerPointApi.Shapes? shapesCollection = null;

                try
                {
                    slide = slidesCollection[i];

                    string slideIndex = slide.SlideIndex.ToString().PadLeft(2, '0');
                    var sb = new StringBuilder();

                    shapesCollection = slide.Shapes;
                    int shapesCount = shapesCollection.Count;

                    for (int j = 1; j <= shapesCount; j++)
                    {
                        NetOffice.PowerPointApi.Shape? shape = null;
                        try
                        {
                            shape = shapesCollection[j];
                            CollectShapeText(shape, sb);
                        }
                        finally
                        {
                            try { shape?.Dispose(); } catch { }
                        }
                    }

                    var slideImagePath = Path.Combine(imagesDir, $"slide_{slideIndex}.jpg");
                    slide.Export(slideImagePath, "JPG", slideWidth, slideHeight);

                    slidesResult.Add((slideImagePath, sb.ToString()));
                }
                finally
                {
                    try { shapesCollection?.Dispose(); } catch { }
                    try { slide?.Dispose(); } catch { }
                }
            }

            try { slidesCollection?.Dispose(); } catch { }

            var imagesDirName = Path.GetFileName(imagesDir);

            using (var sw = new StreamWriter(dstHtmlPath, false, Encoding.UTF8))
            {
                sw.WriteLine("<!doctype html>");
                sw.WriteLine("<html><head><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">");
                sw.WriteLine("<title>Presentation</title>");
                sw.WriteLine("</head><body>");

                for (int i = 0; i < slidesResult.Count; i++)
                {
                    var s = slidesResult[i];
                    var imgRel = imagesDirName + "/" + Path.GetFileName(s.ImgPath);

                    sw.WriteLine("<section style=\"margin-bottom:40px;\">");
                    sw.WriteLine($"<h2>Slide {i + 1}</h2>");
                    sw.WriteLine($"<img src=\"{imgRel}\" alt=\"slide {i + 1}\" style=\"max-width:100%;height:auto;display:block;\"/>");

                    if (!string.IsNullOrWhiteSpace(s.Text))
                    {
                        var encoded = System.Net.WebUtility.HtmlEncode(s.Text).Replace("\n", "<br>");
                        sw.WriteLine($"<div>{encoded}</div>");
                    }

                    sw.WriteLine("</section>");
                }

                sw.WriteLine("</body></html>");
                sw.Flush();
            }
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
    
    private static void CollectShapeText(NetOffice.PowerPointApi.Shape shape, StringBuilder sb)
    {
        try
        {
            if (shape.HasSmartArt == MsoTriState.msoTrue)
            {
                NetOffice.OfficeApi.SmartArt? smartArt = null;
                NetOffice.OfficeApi.SmartArtNodes? nodes = null;

                try
                {
                    smartArt = shape.SmartArt;
                    nodes = smartArt.AllNodes;

                    foreach (NetOffice.OfficeApi.SmartArtNode node in nodes)
                    {
                        try
                        {
                            var text = node.TextFrame2.TextRange.Text;
                            if (!string.IsNullOrWhiteSpace(text))
                                sb.AppendLine(text);
                        }
                        finally
                        {
                            try { node.Dispose(); } catch { }
                        }
                    }
                }
                finally
                {
                    try { nodes?.Dispose(); } catch { }
                    try { smartArt?.Dispose(); } catch { }
                }
            }
            else if (shape.HasTextFrame == MsoTriState.msoTrue && shape.TextFrame.HasText == MsoTriState.msoTrue)
            {
                try
                {
                    var text = shape.TextFrame.TextRange.Text;
                    if (!string.IsNullOrWhiteSpace(text))
                        sb.AppendLine(text);
                }
                catch {}
            }
            else if (shape.Type == MsoShapeType.msoGroup)
            {
                NetOffice.PowerPointApi.GroupShapes? groupItems = null;

                try
                {
                    groupItems = shape.GroupItems;

                    foreach (NetOffice.PowerPointApi.Shape child in groupItems)
                    {
                        try
                        {
                            CollectShapeText(child, sb);
                        }
                        finally
                        {
                            try { child.Dispose(); } catch { }
                        }
                    }
                }
                finally
                {
                    try { groupItems?.Dispose(); } catch { }
                }
            }
        }
        catch {}
    }
}