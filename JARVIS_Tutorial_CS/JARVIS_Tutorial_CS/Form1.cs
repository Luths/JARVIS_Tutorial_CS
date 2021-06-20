using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Net.NetworkInformation;
using System.Speech;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Speech.Recognition; // for speech recognition
using System.Speech.Synthesis;
using System.Speech.AudioFormat;
using System.Threading;
using System.IO.Ports;

namespace JARVIS_Tutorial_CS
{
    public partial class Form1 : Form
    {

        // Form Declarations
        SpeechSynthesizer ss = new SpeechSynthesizer();
        PromptBuilder pb = new PromptBuilder();
        SpeechRecognitionEngine sre = new SpeechRecognitionEngine();
        Choices clist = new Choices();
        SerialPort ardo = new SerialPort();

        [DllImport("user32")]
        public static extern bool ExitWindowsEx(uint uFlags, uint dwReason);
        [DllImport("PowrProf.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern bool SetSuspendState(bool hiberate, bool forceCritical, bool disableWakeEvent);


        private SpeechRecognitionEngine engine;
        private SpeechSynthesizer synthesizer;



        public void Load_Speech()
        {
            clist.Add(new string[] { "hello", "how are you", "what is the time", "open google", "thank you", "close", " weather for durban", "what is the weather", "which languages do you speak?", "open youtube","whats your name"});
            Grammar gr = new Grammar(new GrammarBuilder(clist));
            // ardo = new SerialPort();
            ardo.PortName = "COM6";
            ardo.BaudRate = 9600;

            // Configure the audio input.
            sre.SetInputToDefaultAudioDevice();
            // Configure the audio output. 
            ss.SetOutputToDefaultAudioDevice();

            //See the Voices Installed
            // seeInstalledVoices(ss);

            // Set a Voice
           // ss.SelectVoiceByHints(VoiceGender.Female, VoiceAge.Adult);
            ss.SelectVoiceByHints(VoiceGender.Male, VoiceAge.Adult);

            try
            {
                // Speak a string.
                ss.Speak("Hello , I'm you're Smart Assistant.");
                ss.Speak("Hope you are doing fine today, How may I assist you?");
                sre.RequestRecognizerUpdate();
                sre.LoadGrammar(gr);
                sre.SpeechRecognized += sre_SpeechRecognized;
                sre.RecognizeAsync(RecognizeMode.Multiple);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }



        }

        void sre_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            //If Selected Functions Keywords were Recognized
            switch (e.Result.Text.ToString())
            {
                case "hello":
                    {
                        ss.SpeakAsync("hello there ");
                        break;
                    }
                case "how are you":
                    {
                        ss.SpeakAsync("I'm doing Great, thank you");
                        break;
                    }
                case "What is today’s date and time":
                    {
                        ss.SpeakAsync("Current Time is" + DateTime.Now.ToLongTimeString());
                        break;
                    }


                case "open google":
                    {
                        ss.SpeakAsync("Yes rightaway ");
                        Process.Start("torch.exe", "http:\\www.google.com");
                        //torch -> the type of browser you wish to use.
                        break;
                    }

                case "open youtube":
                    {
                        ss.SpeakAsync("Yes rightaway ");
                        Process.Start("torch.exe", "http:\\www.youtube.com");
                        break;
                    }
                case "What is your name":
                    {
                        ss.SpeakAsync("Chatbots the name");
                        break;
                    }


                case "weather for durban":
                    {
                        ss.SpeakAsync("Yes rightaway ");
                        Process.Start("torch.exe", "https://www.google.co.za/search?safe=active&ei=lH5PW4OLEuuPgAaxioXIAw&q=weather+durban&oq=weather+durban&gs_l=psy-ab.3..0l10.44661.47511.0.48614.13.11.0.0.0.0.504.1651.3-2j1j1.4.0....0...1.1.64.psy-ab..9.4.1650...0i67k1j0i131i67k1j0i131k1.0.VIZeE4Ay9aE");
                        break;
                    }
                case "thank you":
                    {
                        ss.SpeakAsync("No problem, anytime");
                        break;
                    }
                case "Bye":
                    {
                        ss.SpeakAsync("Okay, Bye");
                        Application.Exit();
                        break;
                    }

                case "What is the weather like outside":
                    {
                        ss.SpeakAsync("It's warm but a bit chilly outside");
                        break;
                    }
                case "which languages do you speak":
                    {
                        ss.SpeakAsync("Presently I only understand and speak English.");
                        break;
                    }
                    //case "Log Off the computer":
                    //    {
                    //        ss.SpeakAsync("Logging off the System");
                    //        ExitWindowsEx(0, 0);
                    //        break;
                    //    }
                    //case "Put the computer to sleep":
                    //    {
                    //        ss.SpeakAsync("Putting the computer to sleep");
                    //        SetSuspendState(false, true, true);
                    //        break;
                    //    }
                    //case "shut down the pc":
                    //    {
                    //        ss.SpeakAsync("Doing a System Shut Down");
                    //        Application.Exit();
                    //        Process.Start("shutdown", "/s /t 0");
                    //        break;
                    //    }
                    //case "Restart the computer":
                    //    {
                    //        ss.SpeakAsync("Doing a System Restart");
                    //        Application.Exit();
                    //        Process.Start("shutdown", "/r /t 0");
                    //        break;
                    //    }
            }
            textBox1.Text = e.Result.Text.ToString() + Environment.NewLine;

        }
        public Form1()
        {
            InitializeComponent();
            Load_Speech();
        }

        public void Form1_Load()
        {
            Load_Speech();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        //        //public void seeInstalledVoices(SpeechSynthesizer synth)
        //        //{
        //        //    // Output information about all of the installed voices. 
        //        //    textBox1.Text = "Installed voices -";
        //        //    foreach (InstalledVoice voice in synth.GetInstalledVoices())
        //        //    {
        //        //        VoiceInfo info = voice.VoiceInfo;
        //        //        string AudioFormats = "";
        //        //        foreach (SpeechAudioFormatInfo fmt in info.SupportedAudioFormats)
        //        //        {
        //        //            AudioFormats += String.Format("{0}\n",
        //        //            fmt.EncodingFormat.ToString());
        //        //        }

        //        //        textBox1.AppendText(" Name:          " + info.Name);
        //        //        textBox1.AppendText(" Culture:       " + info.Culture);
        //        //        textBox1.AppendText(" Age:           " + info.Age);
        //        //        textBox1.AppendText(" Gender:        " + info.Gender);
        //        //        textBox1.AppendText(" Description:   " + info.Description);
        //        //        textBox1.AppendText(" ID:            " + info.Id);
        //        //        textBox1.AppendText(" Enabled:       " + voice.Enabled);
        //        //        if (info.SupportedAudioFormats.Count != 0)
        //        //        {
        //        //            textBox1.AppendText(" Audio formats: " + AudioFormats);
        //        //        }
        //        //        else
        //        //        {
        //        //            textBox1.AppendText(" No supported audio formats found");
        //        //        }

        //        //        string AdditionalInfo = "";
        //        //        foreach (string key in info.AdditionalInfo.Keys)
        //        //        {
        //        //            AdditionalInfo += String.Format("  {0}: {1}\n", key, info.AdditionalInfo[key]);
        //        //        }

        //        //        textBox1.AppendText(" Additional Info - " + AdditionalInfo);
        //        //        textBox1.AppendText("/n");
        //        //    }
        //        //}
    }

    



    
}
