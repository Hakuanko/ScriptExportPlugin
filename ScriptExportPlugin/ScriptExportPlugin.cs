using System.Windows;
using System.Windows.Controls;
using YukkuriMovieMaker.Plugin;

namespace ScriptExportPlugin;

/// <summary>
/// ゆっくりMovieMaker4 (YMM4) 向け台本出力プラグイン。
/// ファイルメニューの「動画出力」直下に「台本をテキストで出力」メニュー項目を追加し、
/// タイムライン上のボイスアイテムを時系列順に YMM4 台本 TXT 形式で保存する。
/// </summary>
public class ScriptExportPlugin : IPlugin
{
    public string Name => "台本出力プラグイン";

    // ---- 初期化 --------------------------------------------------------

    /// <summary>
    /// プラグイン読み込み時にメニュー項目を挿入する。
    /// コンストラクタでは UI スレッドが未確立の場合があるため BeginInvoke で遅延実行する。
    /// </summary>
    static ScriptExportPlugin()
    {
        Application.Current.Dispatcher.BeginInvoke(InjectMenuItemSafe);
    }

    private static void InjectMenuItemSafe()
    {
        try
        {
            MenuInjector.Inject(OnExportMenuItemClicked);
        }
        catch (Exception ex)
        {
            ShowError("メニュー初期化失敗", ex);
        }
    }

    // ---- メニュークリックハンドラ ----------------------------------------

    private static void OnExportMenuItemClicked(object sender, RoutedEventArgs e)
    {
        try
        {
            var voiceLines = VoiceLineCollector.Collect();

            if (voiceLines.Count == 0)
            {
                MessageBox.Show(
                    "ボイスアイテムが見つかりませんでした。\nプロジェクトが開かれているか確認してください。",
                    "台本出力",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            ScriptFileWriter.Save(voiceLines);
        }
        catch (Exception ex)
        {
            ShowError("台本出力エラー", ex);
        }
    }

    // ---- ユーティリティ --------------------------------------------------

    private static void ShowError(string title, Exception ex)
        => MessageBox.Show($"{ex.Message}\n\n{ex.StackTrace}", title, MessageBoxButton.OK, MessageBoxImage.Error);
}
