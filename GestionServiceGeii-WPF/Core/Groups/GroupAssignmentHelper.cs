/*
╔════════════════════════════════════════════════════════════════════════════════╗
║                                                                                ║
║                    ───────────────────────────────────                         ║
║                      © Copyright PB-BZH Concept 2026                           ║
║                    ───────────────────────────────────                         ║
║                                                                                ║
║                 contact : mailto:admin@pb-bzh-concept.fr                       ║
╚════════════════════════════════════════════════════════════════════════════════╝

╔════════════════════════════════════════════════════════════════════════════════╗
║  Auteur : Patrick Bourges - PB-BZH Concept                                     ║
║  Le 24/9/2026 - 00:10
╟────────────────────────────────────────────────────────────────────────────────║
║     Projet Visual Studio Professional 2026 : GestionServiceGeii
╟────────────────────────────────────────────────────────────────────────────────║
║     Version : 1.0.0
╟────────────────────────────────────────────────────────────────────────────────║
║                Visual Studio Professional 2026 - Insiders                      ║
║                ──────────────────────────────────────────                      ║
║  Langage     : C# 14                                                           ║
║  Technologie : .NET 10 / WPF                                                   ║
║  Plateforme  : Windows 10 / Windows 11                                         ║
║  Encodage    : UTF-8                                                           ║
╟────────────────────────────────────────────────────────────────────────────────║
║  Nom de fichier : GroupAssignmentHelper.cs
╚════════════════════════════════════════════════════════════════════════════════╝
*/
namespace GestionServiceGeii.Core.Groups;

internal static class GroupAssignmentHelper {
  internal const string GroupSeparator =
      " / ";

  internal static List<string> ExtractGroups(string value) {
    List<string> groups =
        new();

    if (string.IsNullOrWhiteSpace(value))
      return groups;

    string[] items =
        value.Split(
            new[] { "/" },
            StringSplitOptions.RemoveEmptyEntries
        );

    foreach (string item in items) {
      string group =
          item.Trim();

      if (string.IsNullOrWhiteSpace(group))
        continue;

      if (!groups.Contains(group,StringComparer.OrdinalIgnoreCase))
        groups.Add(group);
    }

    return groups;
  }

  internal static string BuildGroupValue(IEnumerable<string> groups) {
    if (groups == null)
      return string.Empty;

    return string.Join(
        GroupSeparator,
        groups
            .Where(group => !string.IsNullOrWhiteSpace(group))
            .Select(group => group.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
    );
  }

  internal static string MergeGroups(
      string firstValue,
      string secondValue
  ) {
    List<string> groups =
        ExtractGroups(firstValue);

    foreach (string group in ExtractGroups(secondValue)) {
      if (!groups.Contains(group,StringComparer.OrdinalIgnoreCase))
        groups.Add(group);
    }

    return BuildGroupValue(groups);
  }

  internal static string RemoveGroup(
      string value,
      string groupToRemove
  ) {
    if (string.IsNullOrWhiteSpace(groupToRemove))
      return value ?? string.Empty;

    List<string> groups =
        ExtractGroups(value);

    groups.RemoveAll(
        group => string.Equals(
            group,
            groupToRemove,
            StringComparison.OrdinalIgnoreCase
        )
    );

    return BuildGroupValue(groups);
  }

  internal static int CountGroups(string value) {
    return ExtractGroups(value).Count;
  }

  internal static bool ContainsGroup(
      string value,
      string group
  ) {
    if (string.IsNullOrWhiteSpace(group))
      return false;

    return ExtractGroups(value)
        .Any(item => string.Equals(
            item,
            group,
            StringComparison.OrdinalIgnoreCase
        ));
  }
}