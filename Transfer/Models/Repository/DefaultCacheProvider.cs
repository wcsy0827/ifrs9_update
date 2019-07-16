using System;
using System.Runtime.Caching;
using Transfer.Models.Interface;

namespace Transfer.Models.Repository
{
    public class DefaultCacheProvider : ICacheProvider
    {
        public ObjectCache Cache { get { return MemoryCache.Default; } }
       
        private string _key = Controllers.AccountController.CurrentUserInfo.Name;

        public object Get(string key)
        {
            return Cache[key + _key];
        }

        public void Invalidate(string key)
        {
            Cache.Remove(key + _key);
        }

        public bool IsSet(string key)
        {
            return (Cache[key + _key] != null);
        }

        public void Set(string key, object data, int cacheTime = 30)
        {
            CacheItemPolicy policy = new CacheItemPolicy();
            policy.AbsoluteExpiration = DateTime.Now + TimeSpan.FromMinutes(cacheTime);
            Cache.Add(new CacheItem(key + _key, data), policy);
        }
    }
}