using System.Runtime.InteropServices;
using DocumentProcessingLibrary.Processing.Handlers;
using DocumentProcessingLibrary.Processing.Models;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace DocumentProcessingLibrary.Documents.SolidWorks.Handlers;

/// <summary>
/// Обработчик связанных моделей в SolidWorks чертежах
/// </summary>
public class SolidWorksReferencedModelsHandler : BaseDocumentElementHandler<SolidWorksDocumentContext>
{
    public override string HandlerName => "SolidWorksReferencedModels";
    protected override ProcessingResult ProcessElement(SolidWorksDocumentContext context, ProcessingConfiguration config)
    {
        if (!config.Options.ProcessProperties || context.Drawing == null)
            return ProcessingResult.Successful(0, 0);
        var totalMatches = 0;
        var processed = 0;
        try
        {
            var modelPaths = GetReferencedModels(context.Drawing);
            foreach (var modelPath in modelPaths)
            {
                if (!string.IsNullOrEmpty(modelPath) && File.Exists(modelPath))
                {
                    var result = ProcessReferencedModel(context.Application, modelPath, config);
                    totalMatches += result.MatchesFound;
                    processed += result.MatchesProcessed;
                }
            }
            return ProcessingResult.Successful(totalMatches, processed);
        }
        catch (Exception ex)
        {
            return ProcessingResult.Failed($"Ошибка обработки связанных моделей: {ex.Message}");
        }
    }
    private HashSet<string> GetReferencedModels(DrawingDoc drawing)
    {
        var modelPaths = new HashSet<string>();
        View? view = null;
        try
        {
            view = ((View)drawing.GetFirstView()).GetNextView() as View;
            while (view != null)
            {
                try
                {
                    var refModel = view.ReferencedDocument;
                    if (refModel != null)
                    {
                        try
                        {
                            var path = refModel.GetPathName();
                            if (!string.IsNullOrEmpty(path))
                                modelPaths.Add(path);
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(refModel);
                        }
                    }
                    view = view.GetNextView() as View;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Ошибка получения ссылок: {ex.Message}");
                }
            }
            return modelPaths;
        }
        finally
        {
            if (view != null)
                Marshal.ReleaseComObject(view);
        }
    }
    private ProcessingResult ProcessReferencedModel(SldWorks swApp, string modelPath, ProcessingConfiguration config)
    {
        ModelDoc2? model = null;
        try
        {
            int errors = 0, warnings = 0;
            var docType = Path.GetExtension(modelPath).Equals(".sldasm", StringComparison.OrdinalIgnoreCase)
                ? (int)swDocumentTypes_e.swDocASSEMBLY
                : (int)swDocumentTypes_e.swDocPART;
            model = swApp.OpenDoc6(
                modelPath,
                docType,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                "",
                ref errors, ref warnings);
            if (model == null)
                return ProcessingResult.Successful(0, 0);
            var propertiesHandler = new SolidWorksPropertiesHandler();
            var context = new SolidWorksDocumentContext
            {
                Model = model, Application = swApp
            };
            var result = propertiesHandler.Handle(context, config);
            model.ForceRebuild3(true);
            model.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref errors, ref warnings);
            swApp.CloseDoc(modelPath);
            return result;
        }
        catch (Exception ex)
        {
            return ProcessingResult.Failed($"Ошибка обработки модели {modelPath}: {ex.Message}");
        }
        finally
        {
            if (model != null)
                Marshal.ReleaseComObject(model);
        }
    }
}