using System.Collections.ObjectModel;
using SlackNet.Events;

namespace SlackUserRelationshipVisualizer;

public class UserRelation
{
    public readonly string FromUserId;
    public readonly string ToUserId;
    public ReadOnlyCollection<MessageEvent> Messages => _messages.AsReadOnly();

    readonly List<MessageEvent> _messages;
    readonly HashSet<string> _tsSet = new();

    public UserRelation(string fromUserId, string toUserId, MessageEvent message)
    {
        FromUserId = fromUserId;
        ToUserId = toUserId;
        _messages = new List<MessageEvent> { message };
    }

    public void Merge(UserRelation other)
    {
        foreach (var message in other._messages)
        {
            // 同じメッセージは登録しない
            if (_tsSet.Add(message.Ts))
            {
                _messages.Add(message);
            }
        }
    }

    /// <returns>最も多くやり取りが行われているチャンネル名</returns>
    public string FindPrimaryChannelLabel()
    {
        return _messages
            .GroupBy(m => m.Channel)
            .OrderByDescending(g => g.Count())
            .First()
            .Key;
    }
}
