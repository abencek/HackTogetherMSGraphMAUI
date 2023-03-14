namespace HackTogetherMSGraphMAUI;

public partial class App : Application
{
	public App()
	{
		InitializeComponent();

		MainPage = new AppShell();
	}

    protected override Window CreateWindow(IActivationState activationState)
    {
        var window = base.CreateWindow(activationState);

        const int windowWidth = 600;
        const int windowHeight = 700;

        window.Width = windowWidth;
        window.Height = windowHeight;

        return window;

    }

}
