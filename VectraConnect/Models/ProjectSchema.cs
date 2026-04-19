using System;
using System.Collections.Generic;

namespace VectraConnect.Models
{
    /// <summary>
    /// Represents the full exported schema of an MS Project file.
    /// </summary>
    public class ProjectSchema
    {
        public string ProjectName { get; set; }
        public string FilePath    { get; set; }

        /// <summary>
        /// The bare .mpp filename without extension (e.g. "ConstructionPhase2").
        /// Used as the file-based prefix segment of every VectraKey.
        /// </summary>
        public string MppFileName { get; set; }

        public DateTime ExportedAt { get; set; }
        public ProjectSummary Summary { get; set; }
        public List<TaskRecord> Tasks { get; set; } = new List<TaskRecord>();
        public List<ResourceRecord> Resources { get; set; } = new List<ResourceRecord>();
        public List<AssignmentRecord> Assignments { get; set; } = new List<AssignmentRecord>();
    }

    public class ProjectSummary
    {
        public DateTime? StartDate { get; set; }
        public DateTime? FinishDate { get; set; }
        public double PercentComplete { get; set; }
        public string Currency { get; set; }
        public int TotalTasks { get; set; }
        public int TotalResources { get; set; }
    }

    public class TaskRecord
    {
        public int UniqueID { get; set; }
        public int ID { get; set; }
        public string Name { get; set; }
        public int OutlineLevel { get; set; }
        public string OutlineNumber { get; set; }
        public bool IsSummary { get; set; }
        public bool IsMilestone { get; set; }
        public string Duration { get; set; }          // e.g. "5 days"
        public DateTime? Start { get; set; }
        public DateTime? Finish { get; set; }
        public double PercentComplete { get; set; }
        public double BaselineCost { get; set; }
        public double ActualCost { get; set; }
        public double RemainingCost { get; set; }
        public string PredecessorIDs { get; set; }   // comma-separated UniqueIDs
        public string Notes { get; set; }
        public string WBS { get; set; }
        public string ConstraintType { get; set; }
        public DateTime? ConstraintDate { get; set; }
        public bool IsCritical { get; set; }
        public double TotalSlack { get; set; }        // in minutes
    }

    public class ResourceRecord
    {
        public int UniqueID { get; set; }
        public int ID { get; set; }
        public string Name { get; set; }
        public string Type { get; set; }              // Work / Material / Cost
        public string EmailAddress { get; set; }
        public double StandardRate { get; set; }
        public double OvertimeRate { get; set; }
        public double MaxUnits { get; set; }
        public double BaselineCost { get; set; }
        public double ActualCost { get; set; }
    }

    public class AssignmentRecord
    {
        public int TaskUniqueID { get; set; }
        public string TaskName { get; set; }
        public int ResourceUniqueID { get; set; }
        public string ResourceName { get; set; }
        public double Units { get; set; }             // % allocation
        public double Work { get; set; }              // in hours
        public double ActualWork { get; set; }
        public double RemainingWork { get; set; }
        public double Cost { get; set; }
        public double ActualCost { get; set; }
    }
}
