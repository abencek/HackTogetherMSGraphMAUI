using HackTogetherMSGraphMAUI.ViewModels;

namespace HackTogetherMSGraphMAUI;

public partial class MainPage : ContentPage
{
	public MainPage()
	{
		InitializeComponent();

        BindingContext = new MainPageViewModel();
    }
}

