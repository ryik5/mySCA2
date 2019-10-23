namespace ASTA.Classes.AutoUpdating
{
    public class AccountEventArgs
    {
        // Сообщение
        public string Message { get; }

        public AccountEventArgs(string mes)
        {
            Message = mes;
        }
    }

    
}
