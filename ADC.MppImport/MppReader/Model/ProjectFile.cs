using System;
using System.Collections.Generic;
using System.Linq;

namespace ADC.MppImport.MppReader.Model
{
    /// <summary>
    /// Top-level container for all project data read from an MPP file.
    /// Ported from org.mpxj.ProjectFile
    /// </summary>
    public class ProjectFile
    {
        public ProjectProperties ProjectProperties { get; } = new ProjectProperties();
        public List<Task> Tasks { get; } = new List<Task>();
        public List<Resource> Resources { get; } = new List<Resource>();
        public List<ResourceAssignment> Assignments { get; } = new List<ResourceAssignment>();
        public List<ProjectCalendar> Calendars { get; } = new List<ProjectCalendar>();
        public List<Exception> IgnoredErrors { get; } = new List<Exception>();
        public List<string> DiagnosticMessages { get; } = new List<string>();

        public Task GetTaskByUniqueID(int uniqueID)
        {
            return Tasks.FirstOrDefault(t => t.UniqueID == uniqueID);
        }

        public Resource GetResourceByUniqueID(int uniqueID)
        {
            return Resources.FirstOrDefault(r => r.UniqueID == uniqueID);
        }

        public ProjectCalendar GetCalendarByUniqueID(int uniqueID)
        {
            return Calendars.FirstOrDefault(c => c.UniqueID == uniqueID);
        }

        public ProjectCalendar GetDefaultCalendar()
        {
            string defaultName = ProjectProperties.DefaultCalendarName ?? "Standard";
            return Calendars.FirstOrDefault(c => c.Name == defaultName)
                ?? Calendars.FirstOrDefault();
        }

        public void AddIgnoredError(Exception ex)
        {
            IgnoredErrors.Add(ex);
        }

        /// <summary>
        /// Resolve cross-references between tasks, resources, assignments, and calendars.
        /// Called after all data has been read.
        /// </summary>
        public void ResolveReferences()
        {
            var taskMap = new Dictionary<int, Task>();
            foreach (var t in Tasks.Where(t => t.UniqueID.HasValue))
                if (!taskMap.ContainsKey(t.UniqueID.Value)) taskMap[t.UniqueID.Value] = t;
            var resourceMap = new Dictionary<int, Resource>();
            foreach (var r in Resources.Where(r => r.UniqueID.HasValue))
                resourceMap[r.UniqueID.Value] = r; // last wins (real record overwrites null block)
            var calendarMap = new Dictionary<int, ProjectCalendar>();
            foreach (var c in Calendars.Where(c => c.UniqueID.HasValue))
                if (!calendarMap.ContainsKey(c.UniqueID.Value)) calendarMap[c.UniqueID.Value] = c;

            // Resolve task hierarchy
            foreach (var task in Tasks)
            {
                if (task.ParentTaskUniqueID.HasValue &&
                    taskMap.TryGetValue(task.ParentTaskUniqueID.Value, out var parent))
                {
                    task.ParentTask = parent;
                    parent.ChildTasks.Add(task);
                }
            }

            // Resolve relations
            foreach (var task in Tasks)
            {
                foreach (var relation in task.Predecessors)
                {
                    if (taskMap.TryGetValue(relation.SourceTaskUniqueID, out var source))
                        relation.SourceTask = source;
                    if (taskMap.TryGetValue(relation.TargetTaskUniqueID, out var target))
                        relation.TargetTask = target;
                }
            }

            // Resolve assignments
            foreach (var assignment in Assignments)
            {
                if (assignment.TaskUniqueID.HasValue &&
                    taskMap.TryGetValue(assignment.TaskUniqueID.Value, out var task))
                {
                    assignment.Task = task;
                    task.Assignments.Add(assignment);
                }
                if (assignment.ResourceUniqueID.HasValue &&
                    resourceMap.TryGetValue(assignment.ResourceUniqueID.Value, out var resource))
                {
                    assignment.Resource = resource;
                    resource.Assignments.Add(assignment);
                }
            }

            // Resolve resource calendars
            foreach (var resource in Resources)
            {
                if (resource.CalendarUniqueID.HasValue &&
                    calendarMap.TryGetValue(resource.CalendarUniqueID.Value, out var cal))
                {
                    resource.Calendar = cal;
                }
            }

            // Resolve calendar parents
            foreach (var calendar in Calendars)
            {
                if (calendar.ParentCalendarUniqueID.HasValue &&
                    calendarMap.TryGetValue(calendar.ParentCalendarUniqueID.Value, out var parent))
                {
                    calendar.ParentCalendar = parent;
                }
            }

            // Set summary flag
            foreach (var task in Tasks)
            {
                if (!task.Summary.HasValue)
                {
                    task.Summary = task.HasChildTasks || (task.ExternalProject ?? false);
                }
            }
        }
    }
}
