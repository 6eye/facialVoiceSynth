using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Microsoft.Kinect;
using Microsoft.Kinect.Face;
using Newtonsoft.Json;
using Sacknet.KinectFacialRecognition;
using Sacknet.KinectFacialRecognition.KinectFaceModel;
using Sacknet.KinectFacialRecognition.ManagedEigenObject;
using Microsoft.Speech.AudioFormat;
using System.Speech.Synthesis;
using System.Diagnostics;
using Microsoft.Speech.Recognition;
using System.Threading;
using System.Windows.Documents; //for span
using System.Text;
using System.Collections;


namespace Sacknet.KinectFacialRecognitionDemo
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private bool takeTrainingImage = false;
        private KinectFacialRecognitionEngine engine;

        private IRecognitionProcessor activeProcessor;
        private KinectSensor kinectSensor;
        private KinectAudioStream convertStream = null;
        private MainWindowViewModel viewModel = new MainWindowViewModel();

        //Edited fields
        private String lastGreeted = ""; //keep track of who we last greeted, so we don't greet them twice in a row
        //TODO eventually expand this to a list data structure with timeouts that will remove the elemtns from the list
        SpeechSynthesizer synth = new SpeechSynthesizer();

        private List<Span> recognitionSpans;
        private SpeechRecognitionEngine speechEngine = null;
        private Boolean questioned = false;
        private Boolean willUseSpeech = true; //will use speech to text recognition

        private String lastSavedImage;

        /// <summary>
        /// Initializes a new instance of the MainWindow class
        /// </summary>
        public MainWindow()
        {
            this.DataContext = this.viewModel;
            this.viewModel.TrainName = "Face 1";
            this.viewModel.ProcessorType = ProcessorTypes.PCA;
            this.viewModel.PropertyChanged += this.ViewModelPropertyChanged;
            this.viewModel.TrainButtonClicked = new ActionCommand(this.Train);
            this.viewModel.TrainNameEnabled = true;

            this.kinectSensor = KinectSensor.GetDefault();
            this.kinectSensor.Open();

            this.InitializeComponent();
            //this.initializeSpeech(); 

            this.LoadProcessor();

            // Set up voice settings
            this.synth.Volume = 100;  // 0...100
            this.synth.Rate = 3;     // -10...10

            this.synth.SelectVoiceByHints(VoiceGender.Female, VoiceAge.Senior);

            WindowLoaded();
            //this.synth.SpeakAsync("What is your name?");


        }

        [DllImport("gdi32")]
        private static extern int DeleteObject(IntPtr o);

        /// <summary>
        /// Loads a bitmap into a bitmap source
        /// </summary>
        private static BitmapSource LoadBitmap(Bitmap source)
        {
            IntPtr ip = source.GetHbitmap();
            BitmapSource bs = null;
            try
            {
                bs = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(ip,
                   IntPtr.Zero, Int32Rect.Empty,
                   System.Windows.Media.Imaging.BitmapSizeOptions.FromEmptyOptions());
                bs.Freeze();
            }
            finally
            {
                DeleteObject(ip);
            }

            return bs;
        }

        /// <summary>
        /// Raised when a property is changed on the view model
        /// </summary>
        private void ViewModelPropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            switch (e.PropertyName)
            {
                case "ProcessorType":
                    this.LoadProcessor();
                    break;
            }
        }

        /// <summary>
        /// Loads the correct procesor based on the selected radio button
        /// </summary>
        private void LoadProcessor()
        {
            if (this.viewModel.ProcessorType == ProcessorTypes.FaceModel)
                this.activeProcessor = new FaceModelRecognitionProcessor();
            else
                this.activeProcessor = new EigenObjectRecognitionProcessor();

            this.LoadAllTargetFaces();
            this.UpdateTargetFaces();

            if (this.engine == null)
            {
                this.engine = new KinectFacialRecognitionEngine(this.kinectSensor, this.activeProcessor);
                this.engine.RecognitionComplete += this.Engine_RecognitionComplete;
            }

            this.engine.Processors = new List<IRecognitionProcessor> { this.activeProcessor };
        }

        /// <summary>
        /// Handles recognition complete events
        /// </summary>

        private void Engine_RecognitionComplete(object sender, KinectFacialRecognition.RecognitionResult e)
        {
            TrackedFace face = null;

            if (e.Faces != null)
                face = e.Faces.FirstOrDefault();

            using (var processedBitmap = (Bitmap)e.ColorSpaceBitmap.Clone())
            {
                if (face == null)
                {
                    this.viewModel.ReadyForTraining = false;
                }
                else
                {
                    using (var g = Graphics.FromImage(processedBitmap))
                    {
                        var isFmb = this.viewModel.ProcessorType == ProcessorTypes.FaceModel;
                        var rect = face.TrackingResult.FaceRect;
                        var faceOutlineColor = Color.Green;

                        if (isFmb)
                        {
                            if (face.TrackingResult.ConstructedFaceModel == null)
                            {
                                faceOutlineColor = Color.Red;
                                
                                if (face.TrackingResult.BuilderStatus == FaceModelBuilderCollectionStatus.Complete)
                                    faceOutlineColor = Color.Orange;
                            }

                            var scale = (rect.Width + rect.Height) / 6;
                            var midX = rect.X + (rect.Width / 2);
                            var midY = rect.Y + (rect.Height / 2);

                            if ((face.TrackingResult.BuilderStatus & FaceModelBuilderCollectionStatus.LeftViewsNeeded) == FaceModelBuilderCollectionStatus.LeftViewsNeeded)
                                g.FillRectangle(new SolidBrush(Color.Red), rect.X - (scale * 2), midY, scale, scale);

                            if ((face.TrackingResult.BuilderStatus & FaceModelBuilderCollectionStatus.RightViewsNeeded) == FaceModelBuilderCollectionStatus.RightViewsNeeded)
                                g.FillRectangle(new SolidBrush(Color.Red), rect.X + rect.Width + (scale * 2), midY, scale, scale);

                            if ((face.TrackingResult.BuilderStatus & FaceModelBuilderCollectionStatus.TiltedUpViewsNeeded) == FaceModelBuilderCollectionStatus.TiltedUpViewsNeeded)
                                g.FillRectangle(new SolidBrush(Color.Red), midX, rect.Y - (scale * 2), scale, scale);

                            if ((face.TrackingResult.BuilderStatus & FaceModelBuilderCollectionStatus.FrontViewFramesNeeded) == FaceModelBuilderCollectionStatus.FrontViewFramesNeeded)
                                g.FillRectangle(new SolidBrush(Color.Red), midX, midY, scale, scale);
                        }

                        this.viewModel.ReadyForTraining = faceOutlineColor == Color.Green;

                        g.DrawPath(new Pen(faceOutlineColor, 5), face.TrackingResult.GetFacePath());

                        //if recognized
                        if (!string.IsNullOrEmpty(face.Key))
                        {
                            var score = Math.Round(face.ProcessorResults.First().Score, 2);
                            ////TODO Greetings, face.Key
                            ////voice to text - https://gist.github.com/elbruno/e4816d4d5a59a3b159eb#file-elbrunokw4v2speech
                            // Write the key on the image...
                            g.DrawString(face.Key/* + ": " + score*/, new Font("Arial", 80), Brushes.Red, new System.Drawing.Point(rect.Left, rect.Top - 25));

                            // Timer handled voice Greeting
                            if (lastGreeted != face.Key) //for now, simply separate greetings by a string
                            {
                                // Async because we don't want synthesizer to block
                                this.synth.SpeakAsync("Hi, " + face.Key);
                                lastGreeted = face.Key;
                            }
                         }

                        //if unrecognized
                        else
                        {
                            if (!questioned) {
                                this.synth.SpeakAsync("What is your name?");
                                questioned = true;
                                }


                        }
                    }

                    if (this.takeTrainingImage)
                    {
                        var eoResult = (EigenObjectRecognitionProcessorResult)face.ProcessorResults.SingleOrDefault(x => x is EigenObjectRecognitionProcessorResult);
                        var fmResult = (FaceModelRecognitionProcessorResult)face.ProcessorResults.SingleOrDefault(x => x is FaceModelRecognitionProcessorResult);

                        var bstf = new BitmapSourceTargetFace();

                        bstf.Key = this.viewModel.TrainName;

                        if (eoResult != null)
                        {
                            bstf.Image = (Bitmap)eoResult.Image.Clone();
                        }
                        else
                        {
                            bstf.Image = face.TrackingResult.GetCroppedFace(e.ColorSpaceBitmap);
                        }

                        if (fmResult != null)
                        {
                            bstf.Deformations = fmResult.Deformations;
                            bstf.HairColor = fmResult.HairColor;
                            bstf.SkinColor = fmResult.SkinColor;
                        }

                        this.viewModel.TargetFaces.Add(bstf);

                        this.SerializeBitmapSourceTargetFace(bstf);

                        this.takeTrainingImage = false;
                        
                        this.UpdateTargetFaces();
                    }
                }

                this.viewModel.CurrentVideoFrame = LoadBitmap(processedBitmap);
            }
            
            // Without an explicit call to GC.Collect here, memory runs out of control :(
            GC.Collect();
        }

        /// <summary>
        /// Saves the target face to disk
        /// </summary>
        private void SerializeBitmapSourceTargetFace(BitmapSourceTargetFace bstf)
        {
            var filenamePrefix = "TF_" + DateTime.Now.Ticks.ToString();
            var suffix = this.viewModel.ProcessorType == ProcessorTypes.FaceModel ? ".fmb" : ".pca";
            System.IO.File.WriteAllText(filenamePrefix + suffix, JsonConvert.SerializeObject(bstf));
            lastSavedImage = filenamePrefix;
            bstf.Image.Save(filenamePrefix + ".png");
        }

        /// <summary>
        /// Loads all BSTFs from the current directory
        /// </summary>
        private void LoadAllTargetFaces()
        {
            this.viewModel.TargetFaces.Clear();
            var result = new List<BitmapSourceTargetFace>();
            var suffix = this.viewModel.ProcessorType == ProcessorTypes.FaceModel ? ".fmb" : ".pca";

            foreach (var file in Directory.GetFiles(".", "TF_*" + suffix))
            {
            //    var dir = Directory.GetCurrentDirectory();
                var bstf = JsonConvert.DeserializeObject<BitmapSourceTargetFace>(File.ReadAllText(file));
                bstf.Image = (Bitmap)Bitmap.FromFile(file.Replace(suffix, ".png"));
                this.viewModel.TargetFaces.Add(bstf);
            }
        }

        /// <summary>
        /// Updates the target faces
        /// </summary>
        private void UpdateTargetFaces()
        {
            if (this.viewModel.TargetFaces.Count > 1)
                this.activeProcessor.SetTargetFaces(this.viewModel.TargetFaces);

            this.viewModel.TrainName = this.viewModel.TrainName.Replace(this.viewModel.TargetFaces.Count.ToString(), (this.viewModel.TargetFaces.Count + 1).ToString());
        }

        /// <summary>
        /// Starts the training image countdown
        /// </summary>
        private void Train()
        {
            if (CanTrain)
            {
                this.viewModel.TrainingInProcess = true;

                var timer = new DispatcherTimer();
                timer.Interval = TimeSpan.FromSeconds(.5);
                timer.Tick += (s2, e2) =>
                {
                    timer.Stop();
                    this.viewModel.TrainingInProcess = false;
                    takeTrainingImage = true;
                };
                timer.Start();
            }


            /*System.Threading.Thread.Sleep(1000);
            var timerTwo = new DispatcherTimer();
            timerTwo.Interval = TimeSpan.FromSeconds(2);
            timerTwo.Tick += (s2, e2) =>
            {
                timerTwo.Stop();
                this.viewModel.TrainingInProcess = false;
                takeTrainingImage = true;
            };
            timerTwo.Start();*/

        }


        private void WindowLoaded()
        {
            // Only one sensor is supported
            this.kinectSensor = KinectSensor.GetDefault();

            if (this.kinectSensor != null)
            {
                // open the sensor
                this.kinectSensor.Open();

                // grab the audio stream
                IReadOnlyList<AudioBeam> audioBeamList = this.kinectSensor.AudioSource.AudioBeams;
                System.IO.Stream audioStream = audioBeamList[0].OpenInputStream();

                // create the convert stream
                this.convertStream = new KinectAudioStream(audioStream);
            }
            else
            {
                // on failure, set the status text
                //this.statusBarText.Text = Properties.Resources.NoKinectReady;
                this.synth.SpeakAsync("No kinect currently ready");
                return;
            }

            RecognizerInfo ri = TryGetKinectRecognizer();

            if (null != ri)
            {
                //this.recognitionSpans = new List<Span> { forwardSpan, backSpan, rightSpan, leftSpan };

                this.speechEngine = new SpeechRecognitionEngine(ri.Id);

                // Create a grammar from grammar definition XML file.
                using (var memoryStream = new MemoryStream(File.ReadAllBytes("..\\..\\..\\SpeechGrammar.xml")))
                {
                    var g = new Grammar(memoryStream);
                    this.speechEngine.LoadGrammar(g);
                }

                this.speechEngine.SpeechRecognized += this.SpeechRecognized;
                //this.speechEngine.SpeechRecognitionRejected += this.SpeechRejected;

                // let the convertStream know speech is going active
                // don't start initially active
                this.convertStream.SpeechActive = true;

                // For long recognition sessions (a few hours or more), it may be beneficial to turn off adaptation of the acoustic model. 
                // This will prevent recognition accuracy from degrading over time.
                ////speechEngine.UpdateRecognizerSetting("AdaptationOn", 0);

                this.speechEngine.SetInputToAudioStream(
                    this.convertStream, new SpeechAudioFormatInfo(EncodingFormat.Pcm, 16000, 16, 1, 32000, 2, null));
                this.speechEngine.RecognizeAsync(RecognizeMode.Multiple);
            }
            else
            {
                //this.statusBarText.Text = Properties.Resources.NoSpeechRecognizer;
                this.synth.SpeakAsync("No speech recognizer found");
            }
        }

        private static RecognizerInfo TryGetKinectRecognizer()
        {
            IEnumerable<RecognizerInfo> recognizers;

            // This is required to catch the case when an expected recognizer is not installed.
            // By default - the x86 Speech Runtime is always expected. 
            try
            {
                recognizers = SpeechRecognitionEngine.InstalledRecognizers();
            }
            catch (COMException)
            {
                return null;
            }

            foreach (RecognizerInfo recognizer in recognizers)
            {
                string value;
                recognizer.AdditionalInfo.TryGetValue("Kinect", out value);
                if ("True".Equals(value, StringComparison.OrdinalIgnoreCase) && "en-US".Equals(recognizer.Culture.Name, StringComparison.OrdinalIgnoreCase))
                {
                    return recognizer;
                }
            }

            return null;
        }

        Boolean CanTrain = true;
        Boolean Speaking = false;

        List<String> list = new List<String>();

        // Speech utterance confidence below which we treat speech as if it hadn't been heard
        double ConfidenceThreshold = 0.7;
        private void SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {


            if (e.Result.Confidence >= ConfidenceThreshold && willUseSpeech)
            {
                string s = e.Result.Semantics.Value.ToString();
                switch (s)
                {
                    case "START":
                    case "BEGIN":
                        this.synth.SpeakAsync("Okay, now automatically classifying");
                        Speaking = true;
                        break;
                    case "OFF":
                        this.synth.SpeakAsync("Okay, no longer automatically classifying");
                        Speaking = false;
                        break;
                    case "HI PENNY":
                        this.synth.SpeakAsync("Hello there");
                        break;
                    case "HOW ARE YOU":
                        this.synth.SpeakAsync("I'm good, and you?");
                        break;
                    case "WHO IS GOING TO TAKE OVER THE WORLD":
                        this.synth.SpeakAsync("No other than the DATA Lab");
                        break;
                    case "WHAT IS YOUR FAVORITE MOVIE":
                        this.synth.SpeakAsync("Terminator");
                        break;
                    case "WINDOWS OR LINUX":
                    case "LINUX OR WINDOWS":
                        this.synth.SpeakAsync("I don't answer stupid questions");
                        break;
                    case "WHAT TIME IS IT":
                        this.synth.SpeakAsync(DateTime.Now.ToString("h:mm:ss tt"));
                        break;
                    case "WHAT MONTH IS IT":
                        this.synth.SpeakAsync(DateTime.Now.ToString("MMMM"));
                        break;
                    case "WHAT DAY IS IT":
                        DateTime dateValue = new DateTime(2008, 6, 11);
                        this.synth.SpeakAsync(dateValue.ToString("dddd"));
                         break;
                    case "WE R":
                        this.synth.SpeakAsync("PENN STATE");
                        break;
                    case "THANK YOU":
                        this.synth.SpeakAsync("You're welcome.");
                        break;
                    case "IRVING":
                        //swallow, this causes nothing but trouble
                        break;
                    case "SHUTDOWN":
                        this.synth.Speak("Farewell.");
                        Application.Current.Shutdown();
                        break;
                    case "IDENTIFY":
                    case "WHO ARE YOU":
                        this.synth.SpeakAsync("I am Penny the collaborative Robot");
                        break;
                    case "CLEAR":
                        this.synth.SpeakAsync("Clearing current identification set");
                        this.viewModel.TargetFaces.Clear();                        
                        lastGreeted = null;
                        list.Clear();
                        this.viewModel.TrainName = "";
                        string startupPath = System.IO.Directory.GetCurrentDirectory();
                        DirectoryInfo di = new DirectoryInfo(startupPath);
                        //remove PCA files
                        FileInfo[] filesPCA = di.GetFiles("*.pca").Where(p => p.Extension == ".pca").ToArray();
                        foreach (FileInfo file in filesPCA)
                            try
                            {
                                file.Attributes = FileAttributes.Normal;
                                File.Delete(file.FullName);
                            }
                            catch { }
                        //remove PNG files
                        FileInfo[] filePNGs = di.GetFiles("TF_*").ToArray();
                        foreach (FileInfo file in filePNGs)
                            try
                            {
                                file.Attributes = FileAttributes.Normal;
                                File.Delete(file.FullName);
                            }
                            catch { }
                        this.UpdateTargetFaces();
                        break;
                    
                    case "INCORRECT":
                        if (lastSavedImage != null) {
                            this.synth.SpeakAsync("I'm sorry, I'll remove that last entry");
                            this.viewModel.TargetFaces.RemoveAt(this.viewModel.TargetFaces.Count - 1);
                            list.Remove(lastGreeted);
                            File.Delete(lastSavedImage + ".PNG");
                            File.Delete(lastSavedImage + ".pca");
                            this.UpdateTargetFaces();
                            this.viewModel.TrainName = "";
                            questioned = true;
                        }
                        break;
                    case "LOWER":
                    case "DECREASE":
                        if (ConfidenceThreshold > .4)
                        {
                            ConfidenceThreshold = ConfidenceThreshold - .1;
                            this.synth.SpeakAsync("Okay, I'll decrease the speech confidence to " + (ConfidenceThreshold * 100).ToString());
                        } else
                        {
                            this.synth.SpeakAsync("Sorry, minimum confidence reached");
                        }
                        break;
                   case "HIGHER":
                   case "INCREASE":
                        if (ConfidenceThreshold < .9)
                        {
                            ConfidenceThreshold = ConfidenceThreshold + .1;
                            this.synth.SpeakAsync("Okay, I'll increase the speech confience to " + (ConfidenceThreshold * 100).ToString());
                        } else
                        {
                            this.synth.SpeakAsync("Sorry, max confidence reached");
                        }
                        break;;
                   case "LEVEL":
                        this.synth.SpeakAsync("Speech confidence level is currently " + (ConfidenceThreshold * 100).ToString());
                        break;
                    default:
                        if (willUseSpeech /*&& s != lastGreeted*/ && CanTrain && questioned)
                        {
                            var match = list.FirstOrDefault(stringToCheck => stringToCheck.Contains(s));
                            if (match == null)
                            {
                                if (Speaking)
                                    this.synth.SpeakAsync("Okay, " + s );//+ ", nice to meet you.");
                            } else {
                                if (Speaking)
                                    this.synth.SpeakAsync("Okay, I'll update your ID set.");
                            }
                                this.viewModel.TrainName = s;
                                if (Speaking)
                                {
                                    list.Add(s);
                                    willUseSpeech = false;
                                    lastGreeted = s;
                                    //this.takeTrainingImage = true;
                                    Train();
                                    CanTrain = false;
                                    willUseSpeech = true;
                                }
                          }
                        //question timer
                        var timer = new DispatcherTimer();
                        timer.Interval = TimeSpan.FromSeconds(10);
                        timer.Tick += (s2, e2) =>
                        {
                            timer.Stop();
                            CanTrain = true;
                            questioned = false;
                        };
                        timer.Start();
                        break;

                    
                }

            }
        }



        /// <summary>
        /// Target face with a BitmapSource accessor for the face
        /// </summary>
        [JsonObject(MemberSerialization.OptIn)]
        public class BitmapSourceTargetFace : IEigenObjectTargetFace, IFaceModelTargetFace
        {
            private BitmapSource bitmapSource;

            /// <summary>
            /// Gets the BitmapSource version of the face
            /// </summary>
            public BitmapSource BitmapSource
            {
                get
                {
                    if (this.bitmapSource == null)
                        this.bitmapSource = MainWindow.LoadBitmap(this.Image);

                    return this.bitmapSource;
                }
            }

            /// <summary>
            /// Gets or sets the key returned when this face is found
            /// </summary>          
            [JsonProperty]
            public string Key { get; set; }

            /// <summary> 
            /// Gets or sets the grayscale, 100x100 target image
            /// </summary>
            public Bitmap Image { get; set; }

            /// <summary>
            /// Gets or sets the detected hair color of the face
            /// </summary>
            [JsonProperty]
            public Color HairColor { get; set; }

            /// <summary>
            /// Gets or sets the detected skin color of the face
            /// </summary>
            [JsonProperty]
            public Color SkinColor { get; set; }

            /// <summary>
            /// Gets or sets the detected deformations of the face
            /// </summary>
            [JsonProperty]
            public IReadOnlyDictionary<FaceShapeDeformations, float> Deformations { get; set; }
        }
    }
}
