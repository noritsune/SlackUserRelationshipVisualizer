using SlackNet.Events;

public class JsonSafeMessage
{
    public string? Channel;
    public string? User;
    public string? Text;
    public string? Subtype;
    public string? Ts;
    public string? ThreadTs;
    public int ReplyCount;

    public static JsonSafeMessage? FromMessageEvent(MessageEvent? messageEvent)
    {
        if (messageEvent == null) return null;
        return new JsonSafeMessage
        {
            Channel = messageEvent.Channel,
            User = messageEvent.User,
            Text = messageEvent.Text,
            Subtype = messageEvent.Subtype,
            Ts = messageEvent.Ts,
            ThreadTs = messageEvent.ThreadTs,
            ReplyCount = messageEvent.ReplyCount,
        };
    }

    public MessageEvent ToMessageEvent()
    {
        return new MessageEvent
        {
            Channel = Channel,
            User = User,
            Text = Text,
            Subtype = Subtype,
            Ts = Ts,
            ThreadTs = ThreadTs,
            ReplyCount = ReplyCount,
        };
    }
}
