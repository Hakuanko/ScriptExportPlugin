using System.IO;
using System.Windows;
using Microsoft.Win32;

namespace ScriptExportPlugin;

/// <summary>
/// 台本行リストをファイルに保存する。
/// SaveFileDialog の表示からファイル書き出しまでを担当する。
/// </summary>
internal static class ScriptFileWriter
{
    /// <summary>
    /// 保存先ダイアログを表示し、ユーザーが選択したパスに台本を書き出す。
    /// </summary>
    /// <param name="scriptLines">YMM4 台本 TXT 形式の各行</param>
    public static void Save(IReadOnlyList<string> scriptLines)
    {
        var savePath = PromptSavePath();
        if (savePath is null) return; // ユーザーがキャンセルした場合

        // 各エントリを \n で結合して書き出す。
        // WriteAllText を使うことでセリフ内の改行文字が
        // エントリ区切りの改行と混在しても正しく保持される。
        File.WriteAllText(savePath, string.Join("\n", scriptLines), System.Text.Encoding.UTF8);

        MessageBox.Show(
            $"台本を出力しました。\n{savePath}\n\n{scriptLines.Count} 行を出力しました。",
            "台本出力完了",
            MessageBoxButton.OK,
            MessageBoxImage.Information);
    }

    // ---- プライベートヘルパー --------------------------------------------

    /// <summary>
    /// 保存ファイルパスをユーザーに選択させる。
    /// キャンセル時は null を返す。
    /// </summary>
    private static string? PromptSavePath()
    {
        var dialog = new SaveFileDialog
        {
            Title      = "台本を保存",
            Filter     = "テキストファイル (*.txt)|*.txt",
            DefaultExt = "txt",
            FileName   = "台本",
        };

        return dialog.ShowDialog() == true ? dialog.FileName : null;
    }
}
