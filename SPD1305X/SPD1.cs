using System;
using System.Windows.Forms;
using NationalInstruments.Visa;
using System.Diagnostics;
using Ivi.Visa;

namespace SPD1305X
{
    public partial class SPD1 : Form
    {
        private IMessageBasedSession session;
        private ResourceManager rmSession;
        private Timer readTimer;
        private Stopwatch stopwatch;
        private TextBox outputTextBox;

        public SPD1()
        {
            InitializeComponent();

            // Initialize timer
            readTimer = new Timer();
            readTimer.Interval = 500; // 500ms interval for reading
            readTimer.Tick += ReadTimer_Tick;

            // Initialize stopwatch
            stopwatch = new Stopwatch();

            // Initialize output TextBox
            outputTextBox = new TextBox();
            outputTextBox.Multiline = true;
            outputTextBox.ScrollBars = ScrollBars.Vertical;
            outputTextBox.Dock = DockStyle.Fill;
            this.Controls.Add(outputTextBox);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                // Create resource manager
                rmSession = new ResourceManager();

                // Find USB instruments
                string[] usbResources = { "USB0::0xF4eC::0x1410::SPD13DCQ7R0719::INSTR" };

                if (usbResources.Length == 0)
                {
                    MessageBox.Show("No USB devices found.", "Connection Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Open the first USB device
                session = (IMessageBasedSession)rmSession.Open(usbResources[0]);

                // Verify device identification
                session.RawIO.Write("*IDN?");
                string response = session.RawIO.ReadString();
                outputTextBox.AppendText($"Connected Device: {response}\r\n");

                // Configure the device for measurement
                // These commands may need to be adjusted based on your specific measurement needs
                session.FormattedIO.WriteLine(":MEAS:VOLT?");  // Query voltage measurement

                // Start the timer and stopwatch
                stopwatch.Restart();
                readTimer.Start();

                // Disable button while running
                button1.Enabled = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Connection Error: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ReadTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                // Check if 10 seconds have elapsed
                if (stopwatch.ElapsedMilliseconds >= 10000)
                {
                    StopMeasurement("Measurement complete");
                    return;
                }

                // Take reading
                session.RawIO.Write(":MEAS:VOLT?");
                string response= session.RawIO.ReadString();
                // Display reading with timestamp
                outputTextBox.AppendText($"Time: {stopwatch.ElapsedMilliseconds}ms - Voltage: {response} V\r\n");
            }
            catch (Exception ex)
            {
                StopMeasurement($"Reading Error: {ex.Message}");
            }
        }

        private void StopMeasurement(string message)
        {
            readTimer.Stop();
            stopwatch.Stop();

            // Close the session
            session?.Dispose();
            rmSession?.Dispose();

            // Re-enable the button
            button1.Enabled = true;

            // Display completion message
            outputTextBox.AppendText(message + "\r\n");
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // Clean up resources
            readTimer?.Stop();
            readTimer?.Dispose();
            session?.Dispose();
            rmSession?.Dispose();

            base.OnFormClosing(e);
        }
    }
}