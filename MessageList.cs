using SlackNet.Events;

public class MessageList
{
    public List<MessageEvent> messages;

    public MessageList(List<MessageEvent> messages)
    {
        this.messages = messages;
    }
}
