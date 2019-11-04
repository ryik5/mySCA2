namespace ASTA.Classes
{
    public class ColorEventArgs
    {
        public System.Drawing.Color Color { get; private set; }

        public ColorEventArgs(System.Drawing.Color color)
        {
            Color = color;
        }
    }
}
