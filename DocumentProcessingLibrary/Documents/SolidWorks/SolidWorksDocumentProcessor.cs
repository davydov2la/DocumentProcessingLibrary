using System.Runtime.InteropServices;
using DocumentProcessingLibrary.Documents.Interfaces;
using DocumentProcessingLibrary.Documents.SolidWorks.Handlers;
using DocumentProcessingLibrary.Processing.Models;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;


namespace DocumentProcessingLibrary.Documents.SolidWorks;

 /// <summary>
/// Процессор для обработки SolidWorks документов
/// </summary>
public class SolidWorksDocumentProcessor : IDocumentProcessor
{
    private SldWorks? _swApp;
    private bool _disposed;
    public string ProcessorName => "SolidWorksDocumentProcessor";
    public IEnumerable<string> SupportedExtensions => new[] { ".slddrw", ".sldprt", ".sldasm" };
    public SolidWorksDocumentProcessor(bool visible = false)
    {
        _swApp = Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application")) as SldWorks;
        if (_swApp == null)
            throw new InvalidOperationException("Не удалось создать экземпляр SolidWorks");
        _swApp.Visible = visible;
    }
    public bool CanProcess(string filePath)
    {
        if (string.IsNullOrEmpty(filePath))
            return false;
        var extension = Path.GetExtension(filePath)?.ToLowerInvariant();
        return extension is ".slddrw" or ".sldprt" or ".sldasm";
    }
    public ProcessingResult Process(DocumentProcessingRequest request)
    {
        if (request == null)
            throw new ArgumentNullException(nameof(request));
        if (_swApp == null)
            return ProcessingResult.Failed("SolidWorks Application не инициализирован");
        if (!File.Exists(request.InputFilePath))
            return ProcessingResult.Failed($"Файл не найден: {request.InputFilePath}");
        if (!CanProcess(request.InputFilePath))
            return ProcessingResult.Failed($"Неподдерживаемый формат файла: {request.InputFilePath}");
        var extension = Path.GetExtension(request.InputFilePath).ToLowerInvariant();
        return extension == ".slddrw"
            ? ProcessDrawing(request)
            : ProcessModel(request);
    }
    private ProcessingResult ProcessDrawing(DocumentProcessingRequest request)
    {
        ModelDoc2? model = null;
        DrawingDoc? drawing = null;
        try
        {
            int errors = 0, warnings = 0;
            model = _swApp!.OpenDoc6(
                request.InputFilePath,
                (int)swDocumentTypes_e.swDocDRAWING,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                "",
                ref errors, ref warnings);
            if (model == null)
                return ProcessingResult.Failed($"Не удалось открыть документ: {request.InputFilePath}");
            drawing = model as DrawingDoc;
            if (drawing == null)
                return ProcessingResult.Failed("Открытый документ не является чертежом");
            var context = new SolidWorksDocumentContext
            {
                Model = model, Drawing = drawing, Application = _swApp!
            };
            var referencedModelsHandler = new SolidWorksReferencedModelsHandler();
            var propertiesHandler = new SolidWorksPropertiesHandler();
            var notesHandler = new SolidWorksNotesHandler();
            var blockNotesHandler = new SolidWorksBlockNotesHandler();
            referencedModelsHandler
                .SetNext(propertiesHandler)
                .SetNext(notesHandler)
                .SetNext(blockNotesHandler);
            var result = referencedModelsHandler.Handle(context, request.Configuration);
            model.ForceRebuild3(true);
            model.GraphicsRedraw2();
            if (request.ExportOptions.SaveModified)
            {
                model.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref errors, ref warnings);
            }
            if (request.ExportOptions.ExportToPdf)
            {
                var modelName = Path.GetFileNameWithoutExtension(request.InputFilePath);
                var pdfFileName = string.IsNullOrEmpty(request.ExportOptions.PdfFileName)
                    ? Path.Combine(request.OutputDirectory, modelName + ".pdf")
                    : request.ExportOptions.PdfFileName;
                SaveDrawingAsPdf(model, drawing, pdfFileName, ref errors, ref warnings);
            }
            _swApp!.CloseDoc(request.InputFilePath);
            return result;
        }
        catch (Exception ex)
        {
            return ProcessingResult.Failed($"Ошибка обработки чертежа: {ex.Message}");
        }
        finally
        {
            if (drawing != null)
            {
                Marshal.ReleaseComObject(drawing);
            }
            if (model != null)
            {
                Marshal.ReleaseComObject(model);
            }
        }
    }
    private ProcessingResult ProcessModel(DocumentProcessingRequest request)
    {
        ModelDoc2? model = null;
        try
        {
            int errors = 0, warnings = 0;
            var extension = Path.GetExtension(request.InputFilePath).ToLowerInvariant();
            var docType = extension == ".sldasm"
                ? (int)swDocumentTypes_e.swDocASSEMBLY
                : (int)swDocumentTypes_e.swDocPART;
            model = _swApp!.OpenDoc6(
                request.InputFilePath,
                docType,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                "",
                ref errors, ref warnings);
            if (model == null)
                return ProcessingResult.Failed($"Не удалось открыть модель: {request.InputFilePath}");
            var context = new SolidWorksDocumentContext
            {
                Model = model, Application = _swApp!
            };
            var propertiesHandler = new SolidWorksPropertiesHandler();
            var result = propertiesHandler.Handle(context, request.Configuration);
            model.ForceRebuild3(true);
            if (request.ExportOptions.SaveModified)
            {
                model.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref errors, ref warnings);
            }
            _swApp!.CloseDoc(request.InputFilePath);
            return result;
        }
        catch (Exception ex)
        {
            return ProcessingResult.Failed($"Ошибка обработки модели: {ex.Message}");
        }
        finally
        {
            if (model != null)
            {
                Marshal.ReleaseComObject(model);
            }
        }
    }
    private void SaveDrawingAsPdf(ModelDoc2 model, DrawingDoc drawing, string fileName, ref int errors, ref int warnings)
    {
        ModelDocExtension? modelExt = null;
        ExportPdfData? exportPdfData = null;
        Sheet? sheet = null;
        try
        {
            modelExt = model.Extension;
            exportPdfData = _swApp!.GetExportFileData((int)swExportDataFileType_e.swExportPdfData) as ExportPdfData;
            if (exportPdfData == null)
                throw new InvalidOperationException("Не удалось получить ExportPdfData");
            var sheetNames = drawing.GetSheetNames() as string[];
            var sheets = new List<DispatchWrapper>();
            if (sheetNames != null)
                foreach (var sheetName in sheetNames)
                {
                    try
                    {
                        drawing.ActivateSheet(sheetName);
                        sheet = drawing.GetCurrentSheet() as Sheet;
                        if (sheet != null)
                        {
                            sheets.Add(new DispatchWrapper(sheet));
                        }
                    }
                    finally
                    {
                        if (sheet != null)
                        {
                            Marshal.ReleaseComObject(sheet);
                            sheet = null;
                        }
                    }
                }

            exportPdfData.SetSheets((int)swExportDataSheetsToExport_e.swExportData_ExportAllSheets, sheets.ToArray());
            exportPdfData.ViewPdfAfterSaving = false;
            modelExt.SaveAs(
                fileName,
                (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                exportPdfData,
                ref errors, ref warnings);
        }
        finally
        {
            if (modelExt != null)
            {
                Marshal.ReleaseComObject(modelExt);
            }
            if (exportPdfData != null)
            {
                Marshal.ReleaseComObject(exportPdfData);
            }
        }
    }
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
    protected virtual void Dispose(bool disposing)
    {
        if (_disposed)
            return;
        if (disposing && _swApp != null)
        {
            try
            {
                _swApp.ExitApp();
            }
            finally
            {
                Marshal.ReleaseComObject(_swApp);
                _swApp = null;
            }
        }
        _disposed = true;
    }
    ~SolidWorksDocumentProcessor()
    {
        Dispose(false);
    }
}