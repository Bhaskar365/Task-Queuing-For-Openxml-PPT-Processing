using SharedModels;

namespace ExcelChartsBlazorOpenxml.Services
{
    public class ButtonProcessingState
    {
        public bool IsProcessing { get; set; }
    }

    public class ReportProcessService
    {
        public ProcessState CurrentState { get; set; } = ProcessState.Idle;
        public string Message { get; set; } = string.Empty;


        public event Action? OnChange;
        public void NotifyStateChanged() => OnChange?.Invoke();

        public void StartProcessing()
        {
            CurrentState = ProcessState.Generating;
            Message = "Generating...";
            NotifyStateChanged();
        }

        public void CompleteProcessing(string msg = "Done")
        {
            CurrentState = ProcessState.Done;
            Message = msg;
            NotifyStateChanged();
        }

        public void Reset()
        {
            CurrentState = ProcessState.Idle;
            Message = string.Empty;
            NotifyStateChanged();
        }
    }
}
