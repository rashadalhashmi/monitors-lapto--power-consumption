using System.Globalization;
using System.Text;
using System.Windows.Forms.DataVisualization.Charting;
using Windows.Devices.Power;

namespace LaptopPowerMonitor;

public sealed class MainForm : Form
{
    private const int MaxDataPoints = 120;
    private const int SampleIntervalMilliseconds = 1000;

    private readonly Battery _battery;
    private readonly System.Windows.Forms.Timer _timer;
    private readonly Chart _powerChart;
    private readonly Label _statusLabel;
    private readonly Button _toggleMonitoringButton;
    private readonly Button _exportButton;
    private readonly List<PowerSample> _samples = [];

    private bool _isMonitoring;
    private int _elapsedSeconds;

    public MainForm()
    {
        _battery = Battery.AggregateBattery;
        _timer = new System.Windows.Forms.Timer { Interval = SampleIntervalMilliseconds };
        _timer.Tick += OnTimerTick;

        Text = "Laptop Power Consumption Monitor";
        MinimumSize = new Size(900, 600);
        StartPosition = FormStartPosition.CenterScreen;
        AutoScaleMode = AutoScaleMode.Font;

        _powerChart = BuildChart();
        _statusLabel = BuildStatusLabel();
        _toggleMonitoringButton = BuildToggleButton();
        _exportButton = BuildExportButton();

        var buttonPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.LeftToRight,
            AutoSize = true,
            WrapContents = false,
            Padding = new Padding(0)
        };

        buttonPanel.Controls.Add(_toggleMonitoringButton);
        buttonPanel.Controls.Add(_exportButton);

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 3,
            Padding = new Padding(12)
        };

        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

        layout.Controls.Add(buttonPanel, 0, 0);
        layout.Controls.Add(_statusLabel, 0, 1);
        layout.Controls.Add(_powerChart, 0, 2);

        Controls.Add(layout);

        UpdateStatus(null, null, "Stopped");
    }

    private static Chart BuildChart()
    {
        var chartArea = new ChartArea("PowerArea")
        {
            BackColor = Color.White
        };

        chartArea.AxisX.Title = "Time (seconds)";
        chartArea.AxisX.Interval = 10;
        chartArea.AxisX.MajorGrid.Enabled = true;
        chartArea.AxisX.MajorGrid.LineColor = Color.Gainsboro;
        chartArea.AxisX.MinorGrid.Enabled = false;
        chartArea.AxisX.LabelStyle.Format = "0";
        chartArea.AxisX.Minimum = 0;
        chartArea.AxisX.IsMarginVisible = false;

        chartArea.AxisY.Title = "Watts";
        chartArea.AxisY.MajorGrid.Enabled = true;
        chartArea.AxisY.MajorGrid.LineColor = Color.Gainsboro;
        chartArea.AxisY.MinorGrid.Enabled = false;
        chartArea.AxisY.LabelStyle.Format = "0.##";

        var series = new Series("Power")
        {
            ChartArea = "PowerArea",
            ChartType = SeriesChartType.Spline,
            BorderWidth = 3,
            Color = Color.DodgerBlue,
            XValueType = ChartValueType.Int32,
            YValueType = ChartValueType.Double
        };

        var chart = new Chart
        {
            Dock = DockStyle.Fill,
            BackColor = Color.White
        };

        chart.ChartAreas.Add(chartArea);
        chart.Series.Add(series);
        chart.Legends.Clear();

        return chart;
    }

    private static Label BuildStatusLabel()
    {
        return new Label
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            Font = new Font("Segoe UI", 10f, FontStyle.Regular, GraphicsUnit.Point),
            Padding = new Padding(0, 8, 0, 8)
        };
    }

    private Button BuildToggleButton()
    {
        var button = new Button
        {
            AutoSize = true,
            Text = "Start Monitoring",
            Padding = new Padding(12, 6, 12, 6)
        };

        button.Click += (_, _) =>
        {
            if (_isMonitoring)
            {
                StopMonitoring();
            }
            else
            {
                StartMonitoring();
            }
        };

        return button;
    }

    private Button BuildExportButton()
    {
        var button = new Button
        {
            AutoSize = true,
            Text = "Export CSV",
            Padding = new Padding(12, 6, 12, 6),
            Enabled = false
        };

        button.Click += (_, _) => ExportSamples();
        return button;
    }

    private void StartMonitoring()
    {
        _isMonitoring = true;
        _timer.Start();
        _toggleMonitoringButton.Text = "Stop Monitoring";
        UpdateStatus(null, null, "Monitoring");
    }

    private void StopMonitoring()
    {
        _timer.Stop();
        _isMonitoring = false;
        _toggleMonitoringButton.Text = "Start Monitoring";
        UpdateStatus(null, null, "Stopped");
    }

    private void OnTimerTick(object? sender, EventArgs e)
    {
        _elapsedSeconds++;

        try
        {
            var report = _battery.GetReport();
            var chargeRateInMilliwatts = report?.ChargeRateInMilliwatts;

            if (chargeRateInMilliwatts is null)
            {
                UpdateStatus(null, null, "Not supported on this device");
                return;
            }

            var watts = chargeRateInMilliwatts.Value / 1000.0;
            var powerStatus = DetermineStatus(watts);

            AddSample(_elapsedSeconds, chargeRateInMilliwatts.Value, watts);
            UpdateStatus(watts, chargeRateInMilliwatts.Value, powerStatus);
        }
        catch (Exception ex)
        {
            UpdateStatus(null, null, $"Error: {ex.Message}");
        }
    }

    private void AddSample(int elapsedSeconds, int milliwatts, double watts)
    {
        _samples.Add(new PowerSample(elapsedSeconds, milliwatts, watts));
        _exportButton.Enabled = _samples.Count > 0;

        if (_samples.Count > MaxDataPoints)
        {
            _samples.RemoveAt(0);
        }

        var series = _powerChart.Series["Power"];
        series.Points.Clear();

        foreach (var sample in _samples)
        {
            series.Points.AddXY(sample.ElapsedSeconds, sample.Watts);
        }

        var xMinimum = Math.Max(0, _elapsedSeconds - MaxDataPoints + 1);
        var xMaximum = Math.Max(MaxDataPoints - 1, _elapsedSeconds);

        _powerChart.ChartAreas["PowerArea"].AxisX.Minimum = xMinimum;
        _powerChart.ChartAreas["PowerArea"].AxisX.Maximum = xMaximum;

        var minWatts = _samples.Min(static sample => sample.Watts);
        var maxWatts = _samples.Max(static sample => sample.Watts);
        var margin = Math.Max(0.5, (maxWatts - minWatts) * 0.1);

        _powerChart.ChartAreas["PowerArea"].AxisY.Minimum = Math.Floor(minWatts - margin);
        _powerChart.ChartAreas["PowerArea"].AxisY.Maximum = Math.Ceiling(maxWatts + margin);
        _powerChart.Invalidate();
    }

    private void UpdateStatus(double? watts, int? milliwatts, string status)
    {
        var wattsText = watts.HasValue
            ? watts.Value.ToString("0.###", CultureInfo.InvariantCulture)
            : "N/A";

        var rawText = milliwatts.HasValue
            ? milliwatts.Value.ToString(CultureInfo.InvariantCulture)
            : "N/A";

        _statusLabel.Text =
            $"Current Watts: {wattsText} W{Environment.NewLine}" +
            $"Raw Charge Rate: {rawText} mW{Environment.NewLine}" +
            $"Status: {status}{Environment.NewLine}" +
            $"Time Elapsed: {_elapsedSeconds} s";
    }

    private static string DetermineStatus(double watts)
    {
        if (watts > 0)
        {
            return "Charging";
        }

        if (watts < 0)
        {
            return "Discharging";
        }

        return "Idle";
    }

    private void ExportSamples()
    {
        try
        {
            using var dialog = new SaveFileDialog
            {
                Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
                FileName = $"power-samples-{DateTime.Now:yyyyMMdd-HHmmss}.csv",
                RestoreDirectory = true
            };

            if (dialog.ShowDialog(this) != DialogResult.OK)
            {
                return;
            }

            var builder = new StringBuilder();
            builder.AppendLine("ElapsedSeconds,ChargeRateMilliwatts,Watts");

            foreach (var sample in _samples)
            {
                builder.Append(sample.ElapsedSeconds.ToString(CultureInfo.InvariantCulture));
                builder.Append(',');
                builder.Append(sample.ChargeRateMilliwatts.ToString(CultureInfo.InvariantCulture));
                builder.Append(',');
                builder.AppendLine(sample.Watts.ToString("0.###", CultureInfo.InvariantCulture));
            }

            File.WriteAllText(dialog.FileName, builder.ToString(), Encoding.UTF8);
            MessageBox.Show(this, "Data exported successfully.", "Export CSV", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            UpdateStatus(null, null, $"Error: {ex.Message}");
            MessageBox.Show(this, ex.Message, "Export Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        _timer.Stop();
        _timer.Dispose();
        _powerChart.Dispose();
        base.OnFormClosing(e);
    }

    private readonly record struct PowerSample(int ElapsedSeconds, int ChargeRateMilliwatts, double Watts);
}
