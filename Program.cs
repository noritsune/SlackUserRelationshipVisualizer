using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using SlackAPI;
using SlackAPI.RPCMessages;
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
        var slackClient = new SlackTaskClient(slackApiToken);

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
    static async Task<List<Channel>> FetchChannels(SlackTaskClient slackClient)
    {
        var types = new[] { "public_channel", "private_channel" };
        var convListRes = await slackClient.GetConversationsListAsync(limit: 1000, types: types);
        if (!convListRes.ok)
        {
            throw new Exception("Failed to get conversation list.: " + convListRes.error);
        }

        // DMは秘密なのでチャンネルのみを対象とする
        return convListRes.channels
            .Where(c => c.is_channel || c.is_group)
            .ToList();
    }

    static async Task<List<Message>> FetchMessagesInChannels(List<Channel> channels, SlackTaskClient slackClient)
    {
        var oldest = DateTime.UtcNow - TimeSpan.FromDays(DAYS);
        Console.WriteLine($"メッセージ取得期間: {oldest}以降");

        var msgDir = new DirectoryInfo(MSG_BY_CHANNEL_OUT_DIR_PATH);
        if (msgDir.Exists) msgDir.Delete(true);
        msgDir.Create();

        var messages = new List<Message>();
        // 並列で実行する
        // 進捗をコンソールに出力する
        var completedChannelCnt = 0;
        await Task.WhenAll(channels.Select(async channel =>
        {
            var messagesInChannel = await FetchMessagesInChannel(slackClient, channel, oldest);
            messages.AddRange(messagesInChannel);
            completedChannelCnt++;
            Console.WriteLine($"チャンネル数: {completedChannelCnt}/{channels.Count}, メッセージ数: {messages.Count}");

            var filePath = MSG_BY_CHANNEL_OUT_DIR_PATH + channel.name + ".json";
            var messageList = new MessageList(messagesInChannel);
            await File.WriteAllTextAsync(filePath, JsonConvert.SerializeObject(messageList, Formatting.Indented));
        }));

        return messages;
    }

    static async Task<List<Message>> FetchMessagesInChannel(SlackTaskClient slackClient, Channel channel, DateTime oldest)
    {
        Console.WriteLine($"チャンネル: {channel.name}のメッセージを取得開始");
        try
        {
            List<Tuple<string, string>> tupleList = new List<Tuple<string, string>>()
            {
                new("channel", channel.id),
                new("oldest", oldest.ToProperTimeStamp()),
                new("limit", "1000")
            };
            var convHistoryRes = await slackClient
                .APIRequestWithTokenAsync<ConversationsMessageHistory>(tupleList.ToArray());

            if (!convHistoryRes.ok) throw new Exception(convHistoryRes.error);

            Console.WriteLine($"チャンネル: {channel.name}のメッセージを取得成功。{convHistoryRes.messages.Length}件");
            return convHistoryRes.messages.ToList();
        }
        catch
        {
            Console.WriteLine($"チャンネル: {channel.name}のメッセージを取得失敗");
            return new List<Message>();
        }
    }

    static async Task<List<Message>> LoadMessagesFromFile()
    {
        var messages = new List<Message>();
        var files = Directory.GetFiles(MSG_BY_CHANNEL_OUT_DIR_PATH);
        foreach (var file in files)
        {
            var messagesStr = await File.ReadAllTextAsync(file);
            var messageList = JsonConvert.DeserializeObject<MessageList>(messagesStr);
            messages.AddRange(messageList.messages);
        }

        return messages;
    }

    static async Task<List<User>> FetchActiveUsers(SlackTaskClient slackClient)
    {
        var userListRes = await slackClient.GetUserListAsync();
        if (!userListRes.ok)
        {
            throw new Exception("Failed to get user list.: " + userListRes.error);
        }

        // 支払いがアクティブなメンバーのみを調査対象とする
        return userListRes.members
            .Where(u => u is { deleted: false, is_bot: false, IsSlackBot: false, is_restricted: false })
            .ToList();
    }

    static Dictionary<string, Dictionary<string, List<Message>>> BuildUserRelationDict(List<User> activeUsers, List<Message> messages)
    {
        var userRelDict = new Dictionary<string, Dictionary<string, List<Message>>>();
        foreach (var activeUser in activeUsers)
        {
            userRelDict[activeUser.id] = new Dictionary<string, List<Message>>();
        }

        var activeUserIdSet = activeUsers.Select(u => u.id).ToHashSet();
        var nonConvSubtypes = new HashSet<string>
        {
            "channel_join", "channel_leave", "group_join", "group_leave", "message_changed", "message_deleted"
        };
        foreach (var message in messages)
        {
            // 会話でなければスキップ
            if (nonConvSubtypes.Contains(message.subtype)) continue;

            // 送信元が調査対象のユーザーでなければスキップ
            var fromUserId = message.user;
            if (!activeUserIdSet.Contains(fromUserId)) continue;

            // メンション先のユーザーIDを抽出する。複数あることがある
            // 送信先が調査対象のユーザーでなければスキップ
            var toUserIds = Regex.Matches(message.text, "<@([A-Z0-9]+)>")
                .Select(m => m.Groups[1].Value)
                .Where(id => activeUserIdSet.Contains(id))
                .ToList();
            if (toUserIds.Count == 0) continue;

            var toUserToMsgs = BuildToUserToMsgs(toUserIds, message);
            userRelDict[fromUserId] = toUserToMsgs;
        }

        return userRelDict;
    }

    static void SaveUserRelationDictToFile(Dictionary<string, Dictionary<string, List<Message>>> userRelDict,
        List<User> users)
    {
        var dir = new DirectoryInfo(MSG_BY_USER_OUT_DIR_PATH);
        if (dir.Exists) dir.Delete(true);
        dir.Create();

        var idToUser = users.ToDictionary(u => u.id, u => u);
        foreach (var (fromUserId, toUserToMsgs) in userRelDict)
        {
            var fromUser = idToUser[fromUserId];
            var filePath = MSG_BY_USER_OUT_DIR_PATH + fromUser.real_name + ".csv";
            using var sw = new StreamWriter(filePath, false, Encoding.UTF8);
            sw.WriteLine("toUser.name,message.text");
            foreach (var (toUserId, msgs) in toUserToMsgs)
            {
                var toUser = idToUser[toUserId];
                foreach (var msg in msgs)
                {
                    var escapedMsgText = msg.text
                        .Replace(",", "，")
                        .Replace("\r", " ")
                        .Replace("\n", " ");
                    sw.WriteLine($"{toUser.real_name},{escapedMsgText}");
                }
            }
        }
    }

    static Dictionary<string, List<Message>> BuildToUserToMsgs(List<string> toUserIds, Message msg)
    {
        var toUserToMsgs = new Dictionary<string, List<Message>>();
        foreach (var toUserId in toUserIds)
        {
            if (!toUserToMsgs.ContainsKey(toUserId))
            {
                toUserToMsgs[toUserId] = new List<Message>();
            }
            toUserToMsgs[toUserId].Add(msg);
        }
        return toUserToMsgs;
    }

    /// <summary>
    /// csvファイルに結果を出力する
    /// 列は送信者IDをid, 送信者名をname, 送信先idをカンマ区切りでrefs, アイコン画像urlをimageとする
    /// </summary>
    static async Task OutputFiles(Dictionary<string, Dictionary<string, List<Message>>> userRelDict, List<User> activeUsers, int maxMsgCnt)
    {
        var csv = new StringBuilder();
        csv.AppendLine("id,name,refs1,refs2,refs3,image");
        foreach (var (fromId, toIdToMsgs) in userRelDict)
        {
            var fromUser = activeUsers.Find(u => u.id == fromId);
            // 警告がうざいので念のため回避
            if (fromUser == null) continue;

            // メッセージ数の相対的な量によって3段階に分ける
            var toIdsStrength1 = toIdToMsgs
                .Where(kv => kv.Value.Count < maxMsgCnt * 0.33)
                .Select(kv => kv.Key)
                .ToList();
            var toIdsStrength2 = toIdToMsgs
                .Where(kv => kv.Value.Count >= maxMsgCnt * 0.33 && kv.Value.Count < maxMsgCnt * 0.66)
                .Select(kv => kv.Key)
                .ToList();
            var toIdsStrength3 = toIdToMsgs
                .Where(kv => kv.Value.Count >= maxMsgCnt * 0.66)
                .Select(kv => kv.Key)
                .ToList();

            var refs1 = $"\"{string.Join(",", toIdsStrength1)}\"";
            var refs2 = $"\"{string.Join(",", toIdsStrength2)}\"";
            var refs3 = $"\"{string.Join(",", toIdsStrength3)}\"";
            csv.AppendLine($"{fromId},{fromUser.real_name},{refs1},{refs2},{refs3},{fromUser.profile.image_48}");
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
