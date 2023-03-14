using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using HackTogetherMSGraphMAUI.Models;
using HackTogetherMSGraphMAUI.Services;
using System.Collections.ObjectModel;
using System.Diagnostics;

namespace HackTogetherMSGraphMAUI.ViewModels
{
    /// <summary>
    /// View Model for MainPage View
    /// </summary>
    public partial class MainPageViewModel : ObservableObject
    {
        /// <summary>
        /// Instance of <see cref="GraphService"/>
        /// </summary>
        private readonly GraphService _graphService;

        /// <summary>
        /// Command's runnig status
        /// </summary>
        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(IsNotBusy))]
        private bool isBusy;

        public bool IsNotBusy => !IsBusy;

        /// <summary>
        /// Collection of Excel files
        /// </summary>
        public ObservableCollection<ExcelFile> ExcelFiles { get; set; }

        /// <summary>
        /// Default constructor
        /// </summary>
        public MainPageViewModel()
        {
            _graphService = new GraphService();
            ExcelFiles = new ObservableCollection<ExcelFile>();
        }

        /// <summary>
        /// Command updates collection of Excel files
        /// </summary>
        /// <returns>Instance of <see cref="Task"/></returns>
        [RelayCommand]
        private async Task GetFilesAsync()
        {
            if (IsBusy)
                return;

            try
            {
                IsBusy = true;

                var files = await _graphService.GetMyFilesAsync();

                ExcelFiles.Clear();

                if (files.Count > 0)
                {
                    foreach (var file in files)
                        ExcelFiles.Add(file);
                }
                else
                {
                    await Shell.Current.DisplayAlert("Message", "No files found!", "OK");
                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error loading user files: {ex}");
                await Shell.Current.DisplayAlert("Error!", "Error loading files!", "OK");
            }
            finally { 
                IsBusy = false; 
            }
        }

    }
}
