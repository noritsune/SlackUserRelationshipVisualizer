namespace SlackUserRelationshipVisualizer;

public class UserRelationList
{
    readonly Dictionary<(string, string), UserRelation> _userIdsToRelation = new();

    public UserRelationList(List<UserRelation> relations)
    {
        foreach (var rel in relations)
        {
            Register(rel);
        }
    }

    void Register(UserRelation regiRel)
    {
        var key = (regiRel.FromUserId, regiRel.ToUserId);
        if (_userIdsToRelation.TryGetValue(key, out var rel))
        {
            rel.Merge(regiRel);
        }
        else
        {
            _userIdsToRelation.Add(key, regiRel);
        }
    }

    public Dictionary<string, List<UserRelation>> BuildFromUserIdToRelations()
    {
        return _userIdsToRelation
            .GroupBy(kvp => kvp.Key.Item1)
            .ToDictionary(
                g => g.Key,
                g => g.Select(kvp => kvp.Value).ToList()
            );
    }

    public int CalcRelationCnt => _userIdsToRelation.Values.Count;

    public int CalcMaxMessageCnt => _userIdsToRelation.Values.Max(rel => rel.Messages.Count);
}
