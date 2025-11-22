using System.Runtime.InteropServices;
using System.Text;
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
        try
        {
            var dstDir = Path.GetDirectoryName(dstHtmlPath) ?? ".";
            Directory.CreateDirectory(dstDir);
            var imagesDir = Path.Combine(dstDir, Path.GetFileNameWithoutExtension(dstHtmlPath) + ".files");
            Directory.CreateDirectory(imagesDir);

            ppt = new NetOffice.PowerPointApi.Application {};
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

            var slides = new List<(string ImgPath, string Text)>();

            var slidesCollection = pres.Slides;
            int slidesCount = slidesCollection.Count;

            for (int i = 1; i <= slidesCount; i++)
            {
                var slide = slidesCollection[i];

                string slideIndex = slide.SlideIndex.ToString().PadLeft(2, '0');
                var sb = new StringBuilder();

                var shapesCollection = slide.Shapes;
                int shapesCount = shapesCollection.Count;
                for (int j = 1; j <= shapesCount; j++)
                {
                    var shape = shapesCollection[j];
                    CollectShapeText(shape, sb);
                }

                var slideImagePath = Path.Combine(imagesDir, $"slide_{slideIndex}.jpg");
                slide.Export(slideImagePath, "JPG", slideWidth, slideHeight);

                slides.Add((slideImagePath, sb.ToString()));
            }

            var imagesDirName = Path.GetFileName(imagesDir);
            using var sw = new StreamWriter(dstHtmlPath, false, Encoding.UTF8);
            sw.WriteLine("<!doctype html>");
            sw.WriteLine("<html><head><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">");
            sw.WriteLine("<title>Presentation</title>");
            sw.WriteLine("</head><body>");
            for (int i = 0; i < slides.Count; i++)
            {
                var s = slides[i];
                var imgRel = imagesDirName + "/" + Path.GetFileName(s.ImgPath);
                sw.WriteLine($"<section style=\"margin-bottom:40px;\">");
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
    
    private static void CollectShapeText(NetOffice.PowerPointApi.Shape shape, System.Text.StringBuilder sb)
    {
        if (shape.HasSmartArt == MsoTriState.msoTrue)
        {
            var sa = shape.SmartArt;
            foreach (NetOffice.OfficeApi.SmartArtNode node in sa.AllNodes)
            {
                sb.AppendLine(node.TextFrame2.TextRange.Text);
            }
        }
        else if (shape.HasTextFrame == MsoTriState.msoTrue && shape.TextFrame.HasText == MsoTriState.msoTrue)
        {
            sb.AppendLine(shape.TextFrame.TextRange.Text);
        }
        else if (shape.Type == NetOffice.OfficeApi.Enums.MsoShapeType.msoGroup)
        {
            foreach (NetOffice.PowerPointApi.Shape child in shape.GroupItems)
                CollectShapeText(child, sb);
        }
    }
}