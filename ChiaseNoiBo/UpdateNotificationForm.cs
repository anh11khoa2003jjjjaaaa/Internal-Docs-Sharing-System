using DocumentFormat.OpenXml.Drawing;
using Guna.UI2.WinForms;
using System;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ChiaseNoiBo
{
    public partial class UpdateNotificationForm : Form
    {
        private int _elapsed = 0;
        private int _duration = 60;
        private Timer _timer;
        private bool _taskCompleted = false;
        public Exception TaskException { get; set; }
        public Func<Task> WorkToDoAsync { get; set; }
        private const int _minDisplayTime = 3; // Hiển thị tối thiểu 3 giây
        private DateTime _startTime;

        public UpdateNotificationForm()
        {
            InitializeComponent();
        }

        private async void UpdateNotificationForm_Load(object sender, EventArgs e)
        {
            lblMessage.Text = $"Vui lòng chờ... đang cập nhật hệ thống ({_elapsed}/{_duration} giây)";
            guna2ProgressIndicator1.Start();

            // Khởi động timer để hiển thị thời gian đếm ngược
            _timer = new Timer();
            _timer.Interval = 1000; // Cập nhật mỗi giây
            _timer.Tick += Timer_Tick;
            _timer.Start();

            // Kiểm tra và thực hiện công việc thực tế
            if (WorkToDoAsync != null)
            {
                try
                {
                    await WorkToDoAsync.Invoke(); // Chạy công việc bất đồng bộ
                    _taskCompleted = true;             
                }
                catch (Exception ex)
                {
                    TaskException = ex;
                }
                finally
                {
                    CloseFormIfReady();
                }
            }
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            _elapsed++;
            lblMessage.Text = $"Vui lòng chờ... đang cập nhật hệ thống ({_elapsed}/{_duration} giây)";

            // Kiểm tra nếu công việc hoàn tất hoặc hết thời gian
            CloseFormIfReady();
        }

        private void CloseFormIfReady()
        {
            // Kiểm tra nếu task đã hoàn tất hoặc hết thời gian
            if ((_taskCompleted || _elapsed >= _duration) && (DateTime.Now - _startTime).TotalSeconds >= _minDisplayTime)
            {
                _timer.Stop();
                guna2ProgressIndicator1.Stop();
                this.Close();
            }
        }
    }


}
