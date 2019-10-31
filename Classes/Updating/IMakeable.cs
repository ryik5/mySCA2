namespace ASTA.Classes.Updating
{
   public interface IMakeable
    {
        void Make();
        void SetParameters(UpdatingParameters parameters);
        UpdatingParameters GetParameters();
    }
}
