namespace AttributeUpdater
{
    public class XYZ
    {
        public double x { get; private set; }
        public double y { get; private set; }
        public double z { get; private set; }

        public XYZ(double X, double Y, double Z)
        {
            x = X;
            y = Y;
            z = Z;
        }
    }
}