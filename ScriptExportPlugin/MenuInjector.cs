using System.Windows;
using System.Windows.Controls;

namespace ScriptExportPlugin;

/// <summary>
/// YMM4 のファイルメニューに「台本をテキストで出力」項目を挿入する。
/// UI ツリーへのアクセスはここに集約し、プラグイン本体から分離する。
/// </summary>
internal static class MenuInjector
{
    private const string FileMenuHeader    = "ファイル";
    private const string VideoExportHeader = "動画出力";
    private const string ExportItemHeader  = "台本をテキストで出力";

    /// <summary>
    /// メインウィンドウのファイルメニューを探し、「動画出力」の直下に項目を挿入する。
    /// </summary>
    /// <param name="clickHandler">メニュークリック時に呼び出すハンドラ</param>
    /// <exception cref="InvalidOperationException">メニューが見つからない場合</exception>
    public static void Inject(RoutedEventHandler clickHandler)
    {
        var mainWindow = Application.Current.MainWindow
            ?? throw new InvalidOperationException("MainWindow が見つかりません。");

        var menuBar = VisualTreeHelper.FindFirstChild<Menu>(mainWindow)
            ?? throw new InvalidOperationException("メニューバーが見つかりません。");

        var fileMenu = menuBar.Items
            .OfType<MenuItem>()
            .FirstOrDefault(mi => mi.Header?.ToString()?.Contains(FileMenuHeader) == true)
            ?? throw new InvalidOperationException($"「{FileMenuHeader}」メニューが見つかりません。");

        int insertIndex = FindInsertIndex(fileMenu);
        fileMenu.Items.Insert(insertIndex, CreateExportMenuItem(clickHandler));
    }

    // ---- プライベートヘルパー --------------------------------------------

    /// <summary>
    /// 「動画出力」の直後のインデックスを返す。
    /// 見つからない場合はメニュー末尾に追加する。
    /// </summary>
    private static int FindInsertIndex(MenuItem fileMenu)
    {
        for (int i = 0; i < fileMenu.Items.Count; i++)
        {
            if (fileMenu.Items[i] is MenuItem mi
                && mi.Header?.ToString()?.Contains(VideoExportHeader) == true)
            {
                return i + 1;
            }
        }
        return fileMenu.Items.Count;
    }

    private static MenuItem CreateExportMenuItem(RoutedEventHandler clickHandler)
    {
        var item = new MenuItem { Header = ExportItemHeader };
        item.Click += clickHandler;
        return item;
    }
}
