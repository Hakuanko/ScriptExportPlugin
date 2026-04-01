using System.Windows;
using System.Windows.Media;

namespace ScriptExportPlugin;

/// <summary>
/// WPF のビジュアルツリーを再帰的に探索するユーティリティ。
/// </summary>
internal static class VisualTreeHelper
{
    /// <summary>
    /// 指定した親要素の子孫から最初に一致する型の要素を返す。
    /// 見つからない場合は null を返す。
    /// </summary>
    public static T? FindFirstChild<T>(DependencyObject parent) where T : DependencyObject
    {
        int childCount = System.Windows.Media.VisualTreeHelper.GetChildrenCount(parent);
        for (int i = 0; i < childCount; i++)
        {
            var child = System.Windows.Media.VisualTreeHelper.GetChild(parent, i);
            if (child is T match) return match;

            var descendant = FindFirstChild<T>(child);
            if (descendant is not null) return descendant;
        }
        return null;
    }
}
