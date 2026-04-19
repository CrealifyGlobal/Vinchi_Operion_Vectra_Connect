using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.MSProject;
using VectraConnect.Models;

namespace VectraConnect
{
    /// <summary>
    /// Extracts data from the active MS Project document via the MSProject COM object model.
    /// </summary>
    public static class ProjectParser
    {
        public static ProjectSchema Parse(Project project)
        {
            if (project == null)
                throw new ArgumentNullException(nameof(project), "No active project found.");

            var schema = new ProjectSchema
            {
                ProjectName = project.Name,
                FilePath    = project.FullName,
                ExportedAt  = DateTime.Now,
                Summary     = BuildSummary(project),
                Tasks       = ExtractTasks(project),
                Resources   = ExtractResources(project),
                Assignments = ExtractAssignments(project)
            };

            return schema;
        }

        // ── Summary ──────────────────────────────────────────────────────────

        private static ProjectSummary BuildSummary(Project p)
        {
            return new ProjectSummary
            {
                StartDate        = SafeDate(p.ProjectStart),
                FinishDate       = SafeDate(p.ProjectFinish),
                PercentComplete  = p.PercentComplete,
                Currency         = p.CurrencySymbol,
                TotalTasks       = p.Tasks.Count,
                TotalResources   = p.Resources.Count
            };
        }

        // ── Tasks ─────────────────────────────────────────────────────────────

        private static List<TaskRecord> ExtractTasks(Project p)
        {
            var list = new List<TaskRecord>();

            foreach (Task t in p.Tasks)
            {
                if (t == null) continue; // MSP can have null task slots

                list.Add(new TaskRecord
                {
                    UniqueID         = t.UniqueID,
                    ID               = t.ID,
                    Name             = t.Name,
                    OutlineLevel     = t.OutlineLevel,
                    OutlineNumber    = t.OutlineNumber,
                    IsSummary        = t.Summary,
                    IsMilestone      = t.Milestone,
                    Duration         = FormatDuration(t.Duration),
                    Start            = SafeDate(t.Start),
                    Finish           = SafeDate(t.Finish),
                    PercentComplete  = t.PercentComplete,
                    BaselineCost     = t.BaselineCost,
                    ActualCost       = t.ActualCost,
                    RemainingCost    = t.RemainingCost,
                    PredecessorIDs   = GetPredecessorIDs(t),
                    Notes            = t.Notes?.Trim(),
                    WBS              = t.WBS,
                    ConstraintType   = t.ConstraintType.ToString(),
                    ConstraintDate   = SafeDate(t.ConstraintDate),
                    IsCritical       = t.Critical,
                    TotalSlack       = t.TotalSlack  // in minutes (MSP internal unit)
                });
            }

            return list;
        }

        private static string GetPredecessorIDs(Task t)
        {
            var ids = new List<string>();
            foreach (TaskDependency dep in t.TaskDependencies)
            {
                if (dep.From.UniqueID != t.UniqueID)
                    ids.Add(dep.From.UniqueID.ToString());
            }
            return string.Join(",", ids);
        }

        private static string FormatDuration(double durationMinutes)
        {
            // MSProject stores Duration in minutes
            if (durationMinutes <= 0) return "0d";
            double days = durationMinutes / 480.0; // 8hr day = 480 min
            return $"{Math.Round(days, 1)}d";
        }

        // ── Resources ─────────────────────────────────────────────────────────

        private static List<ResourceRecord> ExtractResources(Project p)
        {
            var list = new List<ResourceRecord>();

            foreach (Resource r in p.Resources)
            {
                if (r == null) continue;

                list.Add(new ResourceRecord
                {
                    UniqueID       = r.UniqueID,
                    ID             = r.ID,
                    Name           = r.Name,
                    Type           = r.Type.ToString(),
                    EmailAddress   = r.EMailAddress,
                    StandardRate   = r.StandardRate,
                    OvertimeRate   = r.OvertimeRate,
                    MaxUnits       = r.MaxUnits,
                    BaselineCost   = r.BaselineCost,
                    ActualCost     = r.ActualCost
                });
            }

            return list;
        }

        // ── Assignments ───────────────────────────────────────────────────────

        private static List<AssignmentRecord> ExtractAssignments(Project p)
        {
            var list = new List<AssignmentRecord>();

            foreach (Task t in p.Tasks)
            {
                if (t == null) continue;
                foreach (Assignment a in t.Assignments)
                {
                    list.Add(new AssignmentRecord
                    {
                        TaskUniqueID     = t.UniqueID,
                        TaskName         = t.Name,
                        ResourceUniqueID = a.ResourceUniqueID,
                        ResourceName     = a.ResourceName,
                        Units            = a.Units,
                        Work             = a.Work / 60.0,         // minutes → hours
                        ActualWork       = a.ActualWork / 60.0,
                        RemainingWork    = a.RemainingWork / 60.0,
                        Cost             = a.Cost,
                        ActualCost       = a.ActualCost
                    });
                }
            }

            return list;
        }

        // ── Helpers ───────────────────────────────────────────────────────────

        private static DateTime? SafeDate(object value)
        {
            if (value == null) return null;
            if (value is DateTime dt)
            {
                // MSProject uses 1/1/1984 as "NA" sentinel
                if (dt.Year < 1990 || dt.Year > 2100) return null;
                return dt;
            }
            if (DateTime.TryParse(value.ToString(), out var parsed))
                return parsed;
            return null;
        }
    }
}
