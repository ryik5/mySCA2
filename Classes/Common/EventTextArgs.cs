namespace ASTA.Classes
{
    public class EventTextArgs
    {
        // Сообщение
        public string Message { get; }

        public EventTextArgs(string mes)
        {
            Message = mes;
        }
    }
}
