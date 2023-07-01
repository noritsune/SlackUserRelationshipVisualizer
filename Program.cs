﻿using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using OpenAI_API;
using OpenAI_API.Chat;
using OpenAI_API.Models;
using SlackNet;
using SlackNet.Events;
using SlackNet.WebApi;
using SlackUserRelationshipVisualizer;
using Conversation = SlackNet.Conversation;
using File = System.IO.File;

internal static class Program
{
    // ここは実行環境に合わせて適宜変更する
    const string IN_DIR_PATH = "../../../input/";
    const string OUT_DIR_PATH = "../../../output/";

    // 入力場所
    const string SLACK_API_TOKEN_FILE_PATH = IN_DIR_PATH + "slackApiToken.txt";
    const string OPEN_AI_API_KEY_FILE_PATH = IN_DIR_PATH + "openAiApiKey.txt";
    const string DRAW_IO_OPTION_FILE_PATH = IN_DIR_PATH + "drawIoOption.txt";
    const string PROMPT_BASE_FILE_PATH = IN_DIR_PATH + "fetchRelationWordPrompt.txt";

    // 出力場所
    const string MSG_BY_CHANNEL_OUT_DIR_PATH = OUT_DIR_PATH + "messages/channels/";
    const string MSG_BY_USER_OUT_DIR_PATH = OUT_DIR_PATH + "messages/user/";
    const string REL_FOR_DRAW_IO_CSV_FILE_PATH = OUT_DIR_PATH + "relation_forDrawIo.csv";
    const string REL_PURE_CSV_FILE_PATH = OUT_DIR_PATH + "relation_pure.csv";

    // 何日前までのメッセージを取得するか
    const int DAYS = 7;

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
        var userRelList = BuildUserRelationListAboutActiveUsers(activeUsers, messages);
        Console.WriteLine("メッセージの集計終了");

        // デバッグ用にメンションメッセージ群をユーザー毎にファイルに保存
        SaveUserRelationDictToFile(userRelList, activeUsers);

        Console.WriteLine("結果ファイル出力開始");
        await OutputFiles(userRelList, activeUsers);
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
            var msgsInChannelJsonSafe = messagesInChannel
                .Select(JsonSafeMessage.FromMessageEvent)
                .ToList();
            await File.WriteAllTextAsync(filePath, JsonConvert.SerializeObject(msgsInChannelJsonSafe, Formatting.Indented));
        }));

        return messages;
    }

    static async Task<List<MessageEvent>> FetchMessagesInChannel(SlackApiClient slackClient, Conversation conv, DateTime oldest)
    {
        Console.WriteLine($"チャンネル: {conv.Name}のメッセージを取得開始");
        try
        {
            // ルートメッセージ群を取得
            var rootMsgs = new List<MessageEvent>();
            var sw = new Stopwatch();
            sw.Start();
            Console.WriteLine($"チャンネル: {conv.Name}のルートメッセージを取得開始");
            string cursor = null;
            do {
                var convHistoryRes = await slackClient.Conversations.History(
                    conv.Id, oldestTs: oldest.ToTimestamp(), limit: 1000, cursor: cursor);
                rootMsgs.AddRange(convHistoryRes.Messages);

                cursor = convHistoryRes.ResponseMetadata.NextCursor;
            } while (cursor != null);
            Console.WriteLine($"チャンネル: {conv.Name}のルートメッセージを{rootMsgs.Count}件取得完了: {sw.Elapsed.Seconds}秒");

            // スレッドメッセージ群を取得
            var msgsInThread = new List<MessageEvent>();
            var threadHeadMsgs = rootMsgs.Where(m => m.ReplyCount > 0).ToList();
            // 時間を計測する
            sw.Start();
            Console.WriteLine($"チャンネル: {conv.Name}の{threadHeadMsgs.Count}件のスレッドの返信を取得開始");
            // スレッド内の返信を並列で全て取得する
            await Task.WhenAll(threadHeadMsgs.Select(async rootMsg =>
            {
                var threadHistoryRes = await slackClient.Conversations.Replies(
                    conv.Id, rootMsg.Ts, oldestTs: oldest.ToTimestamp(), limit: 1000);
                var replyMsgs = threadHistoryRes.Messages.Where(m => m.Ts != rootMsg.Ts).ToList();
                msgsInThread.AddRange(replyMsgs);
            }));
            sw.Stop();
            Console.WriteLine($"チャンネル: {conv.Name}の{threadHeadMsgs.Count}件のスレッドの返信を取得完了: {sw.Elapsed.Seconds}秒");

            // ルートメッセージとスレッドメッセージを重複なしで結合したものがチャンネル内の全メッセージとなる
            var msgsInChannel = rootMsgs.Concat(msgsInThread).ToList();
            Console.WriteLine($"チャンネル: {conv.Name}のメッセージを取得成功。{msgsInChannel.Count}件");
            return msgsInChannel;
        }
        catch(Exception e)
        {
            Console.WriteLine($"チャンネル: {conv.Name}のメッセージを取得失敗:" + e.Message);
            return new List<MessageEvent>();
        }
    }

    static async Task<List<MessageEvent>> LoadMessagesFromFile()
    {
        var messagesAll = new List<MessageEvent>();
        var files = Directory.GetFiles(MSG_BY_CHANNEL_OUT_DIR_PATH);
        foreach (var file in files)
        {
            var messagesStr = await File.ReadAllTextAsync(file);
            var msgsInChannelJsonSafe = JsonConvert.DeserializeObject<List<JsonSafeMessage>>(messagesStr);
            if (msgsInChannelJsonSafe is null) continue;

            var msgsInChannel = msgsInChannelJsonSafe
                .Select(m => m.ToMessageEvent())
                .ToList();
            messagesAll.AddRange(msgsInChannel);
        }

        return messagesAll;
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

    static UserRelationList BuildUserRelationListAboutActiveUsers(List<User> activeUsers, List<MessageEvent> messages)
    {
        var threadTsToMsgsInThread = messages
            .Where(m => m.ThreadTs != null)
            .GroupBy(m => m.ThreadTs)
            .ToDictionary(g => g.Key, g => g.ToList());

        var userRelsForAllMsgs = new List<UserRelation>();
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

            if (message.ReplyCount > 0)
            {
                // スレッド内の会話を関係性として登録する
                var msgsInThread = threadTsToMsgsInThread[message.ThreadTs];
                var userRelsInThread = GatherUserRelationsInThread(msgsInThread);
                userRelsForAllMsgs.AddRange(userRelsInThread);
            }

            // メンション先のユーザーIDを抽出する。複数あることがある
            // 送信先が調査対象のユーザーでなければスキップ
            var toUserIds = Regex.Matches(message.Text, "<@([A-Z0-9]+)>")
                .Select(m => m.Groups[1].Value)
                // 自身へのメンションがたまに紛れている
                .Where(id => activeUserIdSet.Contains(id) && id != fromUserId)
                .ToList();
            if (toUserIds.Count == 0) continue;

            // メンション先のユーザーIDを関係性として登録する
            var userRelsInMention = toUserIds
                .Select(toUserId => new UserRelation(fromUserId, toUserId, message))
                .ToList();
            userRelsForAllMsgs.AddRange(userRelsInMention);
        }

        var userRelsAboutActiveUsers = userRelsForAllMsgs
            .Where(rel => activeUserIdSet.Contains(rel.FromUserId) && activeUserIdSet.Contains(rel.ToUserId))
            .ToList();

        return new UserRelationList(userRelsAboutActiveUsers);
    }

    static List<UserRelation> GatherUserRelationsInThread(List<MessageEvent> msgsInThread)
    {
        var msgsInThreadAscSortedByTs = msgsInThread.OrderBy(m => m.Ts).ToList();
        var toUserIds = new HashSet<string>{ msgsInThreadAscSortedByTs.First().User };
        var userRels = new List<UserRelation>();
        foreach (var msg in msgsInThreadAscSortedByTs.Skip(1))
        {
            var fromUserId = msg.User;
            // スレッドに書き込まれたメッセージはそれ以前にメッセージを書き込んだ全てのユーザーに向けたものと解釈する
            toUserIds
                .Where(toUserId => toUserId != fromUserId)
                .Select(toUserId => new UserRelation(fromUserId, toUserId, msg))
                .ToList()
                .ForEach(userRels.Add);

            toUserIds.Add(fromUserId);
        }

        return userRels;
    }

    static void SaveUserRelationDictToFile(UserRelationList userRelList, List<User> users)
    {
        var dir = new DirectoryInfo(MSG_BY_USER_OUT_DIR_PATH);
        if (dir.Exists) dir.Delete(true);
        dir.Create();

        var idToUser = users.ToDictionary(u => u.Id, u => u);
        var fromUserIdToRels = userRelList.BuildFromUserIdToRelations();
        foreach (var (fromUserId, rels) in fromUserIdToRels)
        {
            if (!idToUser.TryGetValue(fromUserId, out var fromUser)) continue;

            // ファイル名に使用できないあらゆる文字を除去する
            var escapedFromUserName = Regex.Replace(fromUser.RealName, "[\\\\/:*?\"<>|]", "");
            var filePath = MSG_BY_USER_OUT_DIR_PATH + escapedFromUserName + ".csv";
            using var sw = new StreamWriter(filePath, false, Encoding.UTF8);
            sw.WriteLine("toUser.name,message.text");
            foreach (var rel in rels)
            {
                if (!idToUser.TryGetValue(rel.ToUserId, out var toUser)) continue;

                foreach (var msg in rel.Messages)
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

    /// <summary>
    /// csvファイルに結果を出力する
    /// 列は送信者IDをid, 送信者名をname, アイコン画像urlをimageとする
    /// </summary>
    static async Task OutputFiles(UserRelationList userRelList, List<User> activeUsers)
    {
        const int STRENGTH_DIV = 5;

        // GPT APIの料金概算用
        var msgCharCntSum = userRelList.BuildFromUserIdToRelations().Values
            .Sum(rels =>
                rels.Sum(rel =>
                    rel.Messages.Sum(msg =>
                        msg.Text.Length
                    )
                )
            );
        Console.WriteLine($"文字数合計: {msgCharCntSum}");

        // 事前データ準備
        var idToUser = activeUsers.ToDictionary(u => u.Id, u => u);
        var relCnt = userRelList.CalcRelationCnt;
        var maxMsgCnt = userRelList.CalcMaxMessageCnt;

        var fromUserIdToRels = userRelList.BuildFromUserIdToRelations()
            .ToDictionary(kv => kv.Key, kv => kv.Value);
        // 繋がりが無いユーザーを先頭に持ってくることでグラフを編集しやすくする
        var activeUsersAscOrderByRelCnt = activeUsers
            .OrderBy(u =>
                fromUserIdToRels.TryGetValue(u.Id, out var toRel) ? toRel.Count : 0
            )
            .ToList();

        var userIdSets = new List<HashSet<string>>();
        for (int i = 0; i < activeUsers.Count; i++)
        {
            for (int j = i + 1; j < activeUsers.Count; j++)
            {
                userIdSets.Add(new HashSet<string>{ activeUsers[i].Id, activeUsers[j].Id });
            }
        }

        var userIdSetAndMsgs = userIdSets
            .Select(userIdSet =>
            {
                var userA = userIdSet.First();
                var userB = userIdSet.Last();
                var msgsAtoB = fromUserIdToRels.TryGetValue(userA, out var rels)
                    ? rels.Where(rel => rel.ToUserId == userB).SelectMany(rel => rel.Messages.Select(msg => msg.Text))
                    : new List<string>();
                var msgsBtoA = fromUserIdToRels.TryGetValue(userB, out rels)
                    ? rels.Where(rel => rel.ToUserId == userA).SelectMany(rel => rel.Messages.Select(msg => msg.Text))
                    : new List<string>();
                var msgs = msgsAtoB.Concat(msgsBtoA).ToList();
                return (UserIdA: userIdSet.First(), UserIdB: userIdSet.Last(), Msgs: msgs);
            })
            .Where(tuple => tuple.Msgs.Count > 0)
            .ToArray();

        var userIdSetKeyToRelWord = new Dictionary<string, string>();
        await Task.WhenAll(userIdSetAndMsgs.Select(async tuple =>
        {
            var key = string.Compare(tuple.UserIdA, tuple.UserIdB, StringComparison.Ordinal) < 0
                ? $"{tuple.UserIdA},{tuple.UserIdB}"
                : $"{tuple.UserIdB},{tuple.UserIdA}";
            try
            {
                var relWord = await FetchRelationWord(tuple.Msgs);
                userIdSetKeyToRelWord.Add(key, relWord);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                userIdSetKeyToRelWord.Add(key, "");
            }
        }));

        var csvBody = new StringBuilder();
        var relColLabels = new List<string>();
        var relColStrengths = new List<int>();
        var userSets = new List<(string fromId, string toId)>();
        foreach (var fromUser in activeUsersAscOrderByRelCnt)
        {
            var fromId = fromUser.Id;
            var rels = fromUserIdToRels.TryGetValue(fromId, out var relsTmp)
                ? relsTmp
                : new List<UserRelation>();

            var relCells = new string[relCnt];
            foreach (var rel in rels)
            {
                var toId = rel.ToUserId;
                userSets.Add((fromId, toId));

                relCells[relColLabels.Count] = toId;
                relColLabels.Add($"{fromId}to{toId}");

                // 相対的なメッセージ数によって関係の強さをSTRENGTH_CLASS_DIV段階に分ける
                var norStrength = rel.Messages.Count / (double)maxMsgCnt;
                // 1 ~ STRENGTH_CLASS_DIVの間をとらせたい
                var strength = Math.Min(STRENGTH_DIV, (int)(norStrength * STRENGTH_DIV) + 1);
                relColStrengths.Add(strength);
            }

            var baseCells = new List<string>() { fromId, fromUser.RealName, fromUser.Profile.Image48 };
            csvBody.AppendLine(string.Join(",", baseCells.Concat(relCells)));
        }
        var baseColLabels = new List<string> { "id", "name", "image" };

        if (!Directory.Exists(OUT_DIR_PATH))
        {
            Directory.CreateDirectory(OUT_DIR_PATH);
        }

        // draw.ioで読み込めるように、オプションを付けて出力する
        var drawIoOptionStr = await File.ReadAllTextAsync(DRAW_IO_OPTION_FILE_PATH);
        var relOptions = new List<string>();
        for (var i = 0; i < relColLabels.Count; i++)
        {
            var userSet = userSets[i];
            var key = string.Compare(userSet.fromId, userSet.toId, StringComparison.Ordinal) < 0
                ? $"{userSet.fromId},{userSet.toId}"
                : $"{userSet.toId},{userSet.fromId}";
            var relWord = userIdSetKeyToRelWord[key];

            var colLabel = relColLabels[i];
            var strength = relColStrengths[i];
            relOptions.Add($"# connect: {{\"from\": \"{colLabel}\", \"to\": \"id\", \"label\": \"{relWord}\", \"style\": \"curved=1;fontSize=20;strokeWidth={strength};\"}}");
        }
        var relOptionsStr = string.Join("\n", relOptions);
        drawIoOptionStr = drawIoOptionStr.Replace("$REL_OPTIONS$", relOptionsStr);

        var csvForDrawIo = new StringBuilder();
        csvForDrawIo.AppendLine(drawIoOptionStr);
        var csvHeader = string.Join(",", baseColLabels.Concat(relColLabels));
        csvForDrawIo.AppendLine(csvHeader);
        csvForDrawIo.Append(csvBody);
        await File.WriteAllTextAsync(REL_FOR_DRAW_IO_CSV_FILE_PATH, csvForDrawIo.ToString());

        // 一応、関係表のみのcsvも出しておく
        await File.WriteAllTextAsync(REL_PURE_CSV_FILE_PATH, csvHeader + "\n" + csvBody);
    }

    static async Task<string> FetchRelationWord(List<string> msgs)
    {
        var apiKey = await File.ReadAllTextAsync(OPEN_AI_API_KEY_FILE_PATH, Encoding.UTF8);
        var client = new OpenAIAPI(apiKey);
        var joinedMsgs = string.Join("\n", msgs);
        // 4000トークン制限のため2300文字以内に切り詰める
        joinedMsgs = joinedMsgs.Substring(0, Math.Min(joinedMsgs.Length, 2300));
        var prompt = (await File.ReadAllTextAsync(PROMPT_BASE_FILE_PATH, Encoding.UTF8))
            .Replace("$MESSAGES$", joinedMsgs);
        var chatMsgs = new List<ChatMessage>() { new (ChatMessageRole.User, prompt) };
        var result = await client.Chat.CreateChatCompletionAsync(chatMsgs, model: Model.ChatGPTTurbo);
        var res = result.Choices.First().Message.Content;
        Console.WriteLine($"関係性の言葉: {res}");

        if (string.IsNullOrEmpty(res))
        {
            return "";
        }

        var words = res.Contains(",")
            ? res.Split(",")
            : res.Split("\n");
        // 最後がユニークである可能性が高い
        return words.Last();
    }
}
