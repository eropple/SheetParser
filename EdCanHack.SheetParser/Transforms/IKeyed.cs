namespace EdCanHack.SheetParser.Transforms
{
    public interface IKeyed<out TKey>
    {
        TKey Key { get; }
    }
}
