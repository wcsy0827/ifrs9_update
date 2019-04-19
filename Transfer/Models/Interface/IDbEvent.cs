namespace Transfer.Models.Interface
{
    public interface IDbEvent
    {
        void Dispose();

        void SaveChange();
    }
}