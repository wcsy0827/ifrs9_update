namespace Transfer.Models.Interface
{
    public interface ICacheProviderForIT
    {
        object Get(string key);

        void Invalidate(string key);

        bool IsSet(string key);

        void Set(string key, object data, int cacheTime = 30);
    }
}