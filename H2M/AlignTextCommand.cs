using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace H2M
{
    public class AlignTextCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            IList<Reference> targetRefs = new List<Reference>();
            uidoc.Selection.SetElementIds(new List<ElementId>());

            try
            {
                // Step 1: Pick reference TextNote
                Reference refRef = uidoc.Selection.PickObject(ObjectType.Element, "Pick reference leader");
                TextNote referenceNote = doc.GetElement(refRef) as TextNote;
                if (referenceNote == null)
                {
                    TaskDialog.Show("Error", "Select a TextNote.");
                    return Result.Failed;
                }

                // Get reference leader end position
                IList<Leader> refLeaders = referenceNote.GetLeaders();
                if (refLeaders.Count == 0)
                {
                    TaskDialog.Show("Error", "Reference TextNote has no leaders.");
                    return Result.Failed;
                }
                Leader refLeader = refLeaders[0]; // Use first leader
                XYZ targetShoulder = refLeader.End; // This is the shoulder/end point

                TaskDialog.Show("Align Leaders", $"Align to shoulder at {targetShoulder}? Click OK.");

                // Step 2: Multi-select targets
                try
                {
                    while (true)
                    {
                        Reference targetRef = uidoc.Selection.PickObject(ObjectType.Element,
                            "Pick leader to align (ESC to finish)");
                        targetRefs.Add(targetRef);
                    }
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    // ESC ends selection
                }

                if (targetRefs.Count == 0) return Result.Cancelled;

                // Step 3: Align all targets
                using (Transaction tx = new Transaction(doc, "Align Leader Shoulders"))
                {
                    tx.Start();
                    foreach (Reference tRef in targetRefs)
                    {
                        TextNote targetNote = doc.GetElement(tRef) as TextNote;
                        if (targetNote != null)
                        {
                            IList<Leader> targetLeaders = targetNote.GetLeaders();
                            if (targetLeaders.Count > 0)
                            {
                                // Modify first leader's end point (shoulder)
                                Leader targetLeader = targetLeaders[0];
                                targetLeader.End = targetShoulder;
                            }
                            else
                            {
                                // Add new leader if none exists (optional)
                                Leader newLeader = targetNote.AddLeader(TextNoteLeaderTypes.TNLT_STRAIGHT_R);
                                newLeader.End = targetShoulder;
                            }
                        }
                    }
                    tx.Commit();
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }

            TaskDialog.Show("Success", $"Aligned {targetRefs.Count} leaders.");
            return Result.Succeeded;
        }
    }
}
