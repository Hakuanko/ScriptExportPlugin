using System.Collections;
using System.Reflection;
using System.Windows;

namespace ScriptExportPlugin;

/// <summary>
/// アクティブなタイムラインからボイスアイテムを収集し、
/// YMM4 台本 TXT 形式の文字列リストとして返す。
///
/// YMM4 の内部 API は公開されていないため、リフレクションを用いてアクセスする。
/// キャッシュした PropertyInfo を再利用することで、反復取得によるパフォーマンス低下を防ぐ。
/// </summary>
internal static class VoiceLineCollector
{
    // リフレクション経由で取得するプロパティをキャッシュ。
    // YMM4 のバージョンアップで内部構造が変わった場合はここを更新する。
    private static PropertyInfo? _activeTimelineViewModelProp;
    private static PropertyInfo? _timelineItemsProp;
    private static PropertyInfo? _sceneItemProp;
    private static PropertyInfo? _serifProp;
    private static PropertyInfo? _startFrameProp;
    private static PropertyInfo? _characterProp;
    private static PropertyInfo? _characterNameProp;
    private static PropertyInfo? _itemDescriptionProp;

    /// <summary>
    /// タイムライン上のボイスアイテムを開始フレーム順に収集し、
    /// エスケープ処理済みの台本行リストを返す。
    /// </summary>
    public static List<string> Collect()
    {
        var mainViewModel = Application.Current.MainWindow?.DataContext;
        if (mainViewModel is null) return [];

        var timelineViewModel = GetActiveTimeline(mainViewModel);
        if (timelineViewModel is null) return [];

        var timelineItems = GetTimelineItems(timelineViewModel);
        if (timelineItems is null) return [];

        var voiceEntries = ExtractVoiceEntries(timelineItems);

        // 開始フレーム昇順で並び替えることで、映像の時系列と一致した台本を生成する
        voiceEntries.Sort((a, b) => a.StartFrame.CompareTo(b.StartFrame));

        return voiceEntries.Select(entry => FormatScriptLine(entry.CharacterName, entry.Serif)).ToList();
    }

    // ---- タイムライン取得 ------------------------------------------------

    private static object? GetActiveTimeline(object mainViewModel)
    {
        _activeTimelineViewModelProp ??= mainViewModel.GetType()
            .GetProperty("ActiveTimelineViewModel", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
        return _activeTimelineViewModelProp?.GetValue(mainViewModel);
    }

    private static IEnumerable? GetTimelineItems(object timelineViewModel)
    {
        _timelineItemsProp ??= timelineViewModel.GetType()
            .GetProperty("Items", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
        return _timelineItemsProp?.GetValue(timelineViewModel) as IEnumerable;
    }

    // ---- ボイスエントリ抽出 ---------------------------------------------

    private record VoiceEntry(int StartFrame, string CharacterName, string Serif);

    private static List<VoiceEntry> ExtractVoiceEntries(IEnumerable timelineItems)
    {
        var entries = new List<VoiceEntry>();

        foreach (var timelineItemViewModel in timelineItems)
        {
            var entry = TryCreateVoiceEntry(timelineItemViewModel);
            if (entry is not null) entries.Add(entry);
        }

        return entries;
    }

    private static VoiceEntry? TryCreateVoiceEntry(object timelineItemViewModel)
    {
        var sceneItem = GetSceneItem(timelineItemViewModel);
        if (sceneItem is null) return null;

        // VoiceItem 以外のアイテム（テキスト、画像など）はスキップ
        if (!sceneItem.GetType().FullName!.Contains("VoiceItem")) return null;

        var serif = GetSerif(sceneItem);
        if (string.IsNullOrWhiteSpace(serif)) return null;

        var startFrame  = GetStartFrame(sceneItem);
        var charName    = GetCharacterName(sceneItem, timelineItemViewModel);

        return new VoiceEntry(startFrame, charName, serif);
    }

    // ---- プロパティ個別取得 ---------------------------------------------

    private static object? GetSceneItem(object timelineItemViewModel)
    {
        _sceneItemProp ??= timelineItemViewModel.GetType()
            .GetProperty("Item", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
        return _sceneItemProp?.GetValue(timelineItemViewModel);
    }

    private static string? GetSerif(object sceneItem)
    {
        _serifProp ??= sceneItem.GetType()
            .GetProperty("Serif", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
        return _serifProp?.GetValue(sceneItem)?.ToString();
    }

    private static int GetStartFrame(object sceneItem)
    {
        _startFrameProp ??= sceneItem.GetType()
            .GetProperty("Frame", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
        return (int?)_startFrameProp?.GetValue(sceneItem) ?? 0;
    }

    /// <summary>
    /// キャラクター名を取得する。
    /// まず SceneItem.Character.Name を試み、取得できない場合は
    /// タイムライン項目の Description 文字列（"キャラ名 / セリフ" 形式）から分割して取得する。
    /// </summary>
    private static string GetCharacterName(object sceneItem, object timelineItemViewModel)
    {
        var name = GetCharacterNameFromCharacterObject(sceneItem);
        if (!string.IsNullOrEmpty(name)) return name;

        name = GetCharacterNameFromDescription(timelineItemViewModel);
        return string.IsNullOrWhiteSpace(name) ? "不明" : name;
    }

    private static string? GetCharacterNameFromCharacterObject(object sceneItem)
    {
        _characterProp ??= sceneItem.GetType()
            .GetProperty("Character", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
        var characterObject = _characterProp?.GetValue(sceneItem);
        if (characterObject is null) return null;

        _characterNameProp ??= characterObject.GetType()
            .GetProperty("Name", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
        return _characterNameProp?.GetValue(characterObject)?.ToString();
    }

    private static string? GetCharacterNameFromDescription(object timelineItemViewModel)
    {
        _itemDescriptionProp ??= timelineItemViewModel.GetType()
            .GetProperty("Description", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
        var description = _itemDescriptionProp?.GetValue(timelineItemViewModel)?.ToString() ?? "";

        // Description は "キャラ名 / セリフ" 形式のため、" / " で分割してキャラ名を取り出す
        var separatorIndex = description.IndexOf(" / ", StringComparison.Ordinal);
        return separatorIndex >= 0 ? description[..separatorIndex].Trim() : null;
    }

    // ---- フォーマット ---------------------------------------------------

    /// <summary>
    /// YMM4 台本 TXT 形式にフォーマットする。
    /// セリフ内の「」は YMM4 の仕様に従い \「\」 にエスケープする。
    /// 参照: https://manjubox.net/ymm4/faq/editing/台本ファイルをもとにボイスアイテムを追加する/
    /// </summary>
    private static string FormatScriptLine(string characterName, string serif)
    {
        var escaped = serif
            .Replace("「", "\\「")
            .Replace("」", "\\」");
        return $"{characterName}「{escaped}」";
    }
}
