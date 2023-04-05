using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using SlackNet;
using SlackNet.Events;
using SlackNet.WebApi;
using File = System.IO.File;

internal static class Program
{
    // ここは実行環境に合わせて適宜変更する
    const string IN_DIR_PATH = "../../../input/";
    const string OUT_DIR_PATH = "../../../output/";

    // 入力場所
    const string SLACK_API_TOKEN_FILE_PATH = IN_DIR_PATH + "slackApiToken.txt";
    const string DRAW_IO_OPTION_FILE_PATH = IN_DIR_PATH + "drawIoOption.txt";

    // 出力場所
    const string MSG_BY_CHANNEL_OUT_DIR_PATH = OUT_DIR_PATH + "messages/channels/";
    const string MSG_BY_USER_OUT_DIR_PATH = OUT_DIR_PATH + "messages/user/";
    const string REL_FOR_DRAW_IO_CSV_FILE_PATH = OUT_DIR_PATH + "relation_forDrawIo.csv";
    const string REL_PURE_CSV_FILE_PATH = OUT_DIR_PATH + "relation_pure.csv";

    // 何日前までのメッセージを取得するか
    const int DAYS = 30;

    // デバッグ用設定
    const bool IS_LOAD_MSG_FROM_FILE = false;

    static async Task Main()
    {
        var slackApiToken = await File.ReadAllTextAsync(SLACK_API_TOKEN_FILE_PATH, Encoding.UTF8);
        var slackClient = new SlackApiClient(slackApiToken);

        var channels = await FetchChannels(slackClient);
        Console.WriteLine($"調査対象のチャンネル: {channels.Count}件");

        // まずは対象期間のメッセージ全てを取得する
        var messages = IS_LOAD_MSG_FROM_FILE
            ? await LoadMessagesFromFile()
            : await FetchMessagesInChannels(channels, slackClient);
        Console.WriteLine($"調査対象のメッセージ: {messages.Count}件");

        // 調査対象となるユーザーを取得する
        var activeUsers = await FetchActiveUsers(slackClient);
        Console.WriteLine($"調査対象のアクティブなメンバー: {activeUsers.Count}人");

        // メンバー間のメッセージを集計する
        Console.WriteLine("メッセージの集計開始");
        var userRelDict = BuildUserRelationDict(activeUsers, messages);
        Console.WriteLine("メッセージの集計終了");

        // デバッグ用にメンションメッセージ群をユーザー毎にファイルに保存
        SaveUserRelationDictToFile(userRelDict, activeUsers);

        var maxMsgCnt = userRelDict
            .SelectMany(u => u.Value.Select(kv => kv.Value.Count))
            .Max();

        Console.WriteLine("結果ファイル出力開始");
        await OutputFiles(userRelDict, activeUsers, maxMsgCnt);
        Console.WriteLine("結果ファイル出力終了");
    }

    /// <summary>
    /// 調査対象のチャンネル一覧を取得する
    /// パブリックチャンネルとプライベートチャンネルを対象とする
    /// </summary>
    static async Task<List<Conversation>> FetchChannels(SlackApiClient slackClient)
    {
        var types = new[] { ConversationType.PublicChannel, ConversationType.PrivateChannel };
        var convListRes = await slackClient.Conversations.List(excludeArchived:true, limit: 1000, types: types);
        return convListRes.Channels.ToList();
    }

    static async Task<List<MessageEvent>> FetchMessagesInChannels(List<Conversation> channels, SlackApiClient slackClient)
    {
        var oldest = DateTime.UtcNow - TimeSpan.FromDays(DAYS);
        Console.WriteLine($"メッセージ取得期間: {oldest}以降");

        var msgDir = new DirectoryInfo(MSG_BY_CHANNEL_OUT_DIR_PATH);
        if (msgDir.Exists) msgDir.Delete(true);
        msgDir.Create();

        var messages = new List<MessageEvent>();
        // 並列で実行する
        // 進捗をコンソールに出力する
        var completedChannelCnt = 0;
        await Task.WhenAll(channels.Select(async conv =>
        {
            var messagesInChannel = await FetchMessagesInChannel(slackClient, conv, oldest);
            messages.AddRange(messagesInChannel);
            completedChannelCnt++;
            Console.WriteLine($"チャンネル数: {completedChannelCnt}/{channels.Count}, メッセージ数: {messages.Count}");

            var filePath = MSG_BY_CHANNEL_OUT_DIR_PATH + conv.Name + ".json";
            var messageList = new MessageList(messagesInChannel);
            await File.WriteAllTextAsync(filePath, JsonConvert.SerializeObject(messageList, Formatting.Indented));
        }));

        return messages;
    }

    static async Task<List<MessageEvent>> FetchMessagesInChannel(SlackApiClient slackClient, Conversation conv, DateTime oldest)
    {
        Console.WriteLine($"チャンネル: {conv.Name}のメッセージを取得開始");
        try
        {
            var convHistoryRes = await slackClient.Conversations.History(conv.Id, oldestTs: oldest.ToTimestamp(), limit: 1000);

            Console.WriteLine($"チャンネル: {conv.Name}のメッセージを取得成功。{convHistoryRes.Messages.Count}件");
            return convHistoryRes.Messages.ToList();
        }
        catch(Exception e)
        {
            Console.WriteLine($"チャンネル: {conv.Name}のメッセージを取得失敗:" + e.Message);
            return new List<MessageEvent>();
        }
    }

    static async Task<List<MessageEvent>> LoadMessagesFromFile()
    {
        var messages = new List<MessageEvent>();
        var files = Directory.GetFiles(MSG_BY_CHANNEL_OUT_DIR_PATH);
        foreach (var file in files)
        {
            var messagesStr = await File.ReadAllTextAsync(file);
            var messageList = JsonConvert.DeserializeObject<MessageList>(messagesStr);
            if (messageList is null) continue;

            messages.AddRange(messageList.messages);
        }

        return messages;
    }

    static async Task<List<User>> FetchActiveUsers(SlackApiClient slackClient)
    {
        var userListRes = await slackClient.Users.List();
        // 支払いがアクティブなメンバーのみを調査対象とする
        return userListRes.Members
            // IsRestricted: falseなら支払いが非アクティブ
            .Where(u => u is { Deleted: false, IsBot: false, IsRestricted: false } && u.Name != "slackbot")
            .ToList();
    }

    static Dictionary<string, Dictionary<string, List<MessageEvent>>> BuildUserRelationDict(List<User> activeUsers, List<MessageEvent> messages)
    {
        var userRelDict = new Dictionary<string, Dictionary<string, List<MessageEvent>>>();
        foreach (var activeUser in activeUsers)
        {
            userRelDict[activeUser.Id] = new Dictionary<string, List<MessageEvent>>();
        }

        var activeUserIdSet = activeUsers.Select(u => u.Id).ToHashSet();
        var nonConvSubtypes = new HashSet<string>
        {
            "channel_join", "channel_leave", "group_join", "group_leave", "message_changed", "message_deleted"
        };
        foreach (var message in messages)
        {
            // 会話でなければスキップ
            if (nonConvSubtypes.Contains(message.Subtype)) continue;

            // 送信元が調査対象のユーザーでなければスキップ
            var fromUserId = message.User;
            if (!activeUserIdSet.Contains(fromUserId)) continue;

            // メンション先のユーザーIDを抽出する。複数あることがある
            // 送信先が調査対象のユーザーでなければスキップ
            var toUserIds = Regex.Matches(message.Text, "<@([A-Z0-9]+)>")
                .Select(m => m.Groups[1].Value)
                .Where(id => activeUserIdSet.Contains(id))
                .ToList();
            if (toUserIds.Count == 0) continue;

            foreach (var toUserId in toUserIds)
            {
                if (!userRelDict[fromUserId].ContainsKey(toUserId))
                {
                    userRelDict[fromUserId][toUserId] = new List<MessageEvent>();
                }
                userRelDict[fromUserId][toUserId].Add(message);
            }
        }

        return userRelDict;
    }

    static void SaveUserRelationDictToFile(Dictionary<string, Dictionary<string, List<MessageEvent>>> userRelDict,
        List<User> users)
    {
        var dir = new DirectoryInfo(MSG_BY_USER_OUT_DIR_PATH);
        if (dir.Exists) dir.Delete(true);
        dir.Create();

        var idToUser = users.ToDictionary(u => u.Id, u => u);
        foreach (var (fromUserId, toUserToMsgs) in userRelDict)
        {
            var fromUser = idToUser[fromUserId];
            var filePath = MSG_BY_USER_OUT_DIR_PATH + fromUser.RealName + ".csv";
            using var sw = new StreamWriter(filePath, false, Encoding.UTF8);
            sw.WriteLine("toUser.name,message.text");
            foreach (var (toUserId, msgs) in toUserToMsgs)
            {
                var toUser = idToUser[toUserId];
                foreach (var msg in msgs)
                {
                    var escapedMsgText = msg.Text
                        .Replace(",", "，")
                        .Replace("\r", " ")
                        .Replace("\n", " ");
                    sw.WriteLine($"{toUser.RealName},{escapedMsgText}");
                }
            }
        }
    }

    static Dictionary<string, List<MessageEvent>> BuildToUserToMsgs(List<string> toUserIds, MessageEvent msg)
    {
        var toUserToMsgs = new Dictionary<string, List<MessageEvent>>();
        foreach (var toUserId in toUserIds)
        {
            if (!toUserToMsgs.ContainsKey(toUserId))
            {
                toUserToMsgs[toUserId] = new List<MessageEvent>();
            }
            toUserToMsgs[toUserId].Add(msg);
        }
        return toUserToMsgs;
    }

    /// <summary>
    /// csvファイルに結果を出力する
    /// 列は送信者IDをid, 送信者名をname, 送信先idをカンマ区切りでrefs, アイコン画像urlをimageとする
    /// </summary>
    static async Task OutputFiles(Dictionary<string, Dictionary<string, List<MessageEvent>>> userRelDict, List<User> activeUsers, int maxMsgCnt)
    {
        const int STRENGTH_CLASS_DIV = 5;

        var csv = new StringBuilder();
        csv.AppendLine("id,name,refs1,refs2,refs3,refs4,refs5,image");
        foreach (var (fromId, toIdToMsgs) in userRelDict)
        {
            var fromUser = activeUsers.Find(u => u.Id == fromId);
            // 警告がうざいので念のため回避
            if (fromUser == null) continue;

            // 相対的なメッセージ数によって関係の強さをSTRENGTH_CLASS_DIV段階に分ける
            var toIdsStrI = new List<List<string>>();
            for (var i = 0; i < STRENGTH_CLASS_DIV; i++)
            {
                toIdsStrI.Add(new List<string>());
            }
            foreach (var (toId, msgs) in toIdToMsgs)
            {
                var strength = msgs.Count / (double)maxMsgCnt;
                var strengthClass = Math.Min(4, (int)(strength * STRENGTH_CLASS_DIV));
                toIdsStrI[strengthClass].Add(toId);
            }

            var refsStr = toIdsStrI
                .Select(toIds => $"\"{string.Join(",", toIds)}\"")
                .ToList();

            csv.AppendLine($"{fromId},{fromUser.RealName},{refsStr[0]},{refsStr[1]},{refsStr[2]},{refsStr[3]},{refsStr[4]},{fromUser.Profile.Image48}");
        }

        if (!Directory.Exists(OUT_DIR_PATH))
        {
            Directory.CreateDirectory(OUT_DIR_PATH);
        }

        // draw.ioで読み込めるように、オプションを付けて出力する
        var csvForDrawIo = new StringBuilder();
        var drawIoOptionStr = await File.ReadAllTextAsync(DRAW_IO_OPTION_FILE_PATH);
        csvForDrawIo.Append(drawIoOptionStr);
        csvForDrawIo.Append(csv.ToString());
        await File.WriteAllTextAsync(REL_FOR_DRAW_IO_CSV_FILE_PATH, csvForDrawIo.ToString());

        // 一応、関係表のみのcsvも出しておく
        await File.WriteAllTextAsync(REL_PURE_CSV_FILE_PATH, csv.ToString());
    }
}
