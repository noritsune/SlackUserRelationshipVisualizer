using SlackNet.Events;

public class JsonSafeMessage
{
    public string User;
    public string Text;
    public string Subtype;

    public static JsonSafeMessage FromMessageEvent(MessageEvent messageEvent)
    {
        return new JsonSafeMessage
        {
            User = messageEvent.User,
            Text = messageEvent.Text,
            Subtype = messageEvent.Subtype,
        };
    }

    public MessageEvent ToMessageEvent()
    {
        return new MessageEvent
        {
            User = User,
            Text = Text,
            Subtype = Subtype,
        };
    }
}
