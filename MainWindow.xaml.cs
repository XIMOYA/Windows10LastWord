using Microsoft.UI;
using Microsoft.UI.Dispatching;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Media.Animation;
using Microsoft.UI.Xaml.Media.Imaging;
using Microsoft.UI.Xaml.Shapes;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Windows.Foundation;
using Windows.Media.Core;
using Windows.Media.Playback;
using Windows.UI;

namespace Windows10LastWord
{
    public class FarewellStep
    {
        public string Message { get; set; } = "";
        public string? ActionName { get; set; }
        public string? ActionCommand { get; set; }
        public string? BackgroundImage { get; set; }
    }

    public sealed partial class MainWindow : Window
    {
        private Queue<FarewellStep>? farewellSteps;
        private int totalSteps = 0;
        private string win11InstallerPath = "";
        private Storyboard? _shutdownStoryboard;
        private MediaPlayer _backgroundMusicPlayer;
        private MediaPlayer _notificationSoundPlayer;
        private List<string> _notificationSounds;
        private int _soundIndex = 0;

        private const double BGM_NORMAL_VOLUME = 0.5;
        private const double BGM_DUCKING_VOLUME = 0.1;

        public MainWindow()
        {
            this.InitializeComponent();
            this.Title = "Windows 10 最后的留言";

            _backgroundMusicPlayer = new MediaPlayer();
            _notificationSoundPlayer = new MediaPlayer();

            _notificationSounds = new List<string>
            {
                "1 (1).wav", "1 (2).wav", "1 (3).wav", "1 (4).wav", "1 (5).wav", "1 (6).wav",
                "1 (7).wav", "1 (8).wav", "1 (9).wav", "1 (10).wav", "1 (11).wav", "1 (12).wav",
                "1 (13).wav", "1 (14).wav", "1 (15).wav", "1 (16).wav", "1 (17).wav", "1 (18).wav",
                "1 (19).wav", "1 (20).wav", "1 (21).wav", "1 (22).wav", "1 (23).wav", "1 (24).wav",
                "1 (25).wav", "1 (26).wav", "1 (27).wav", "1 (28).wav", "1 (29).wav", "1 (30).wav",
                "1 (31).wav", "1 (32).wav", "1 (33).wav", "1 (34).wav", "1 (35).wav", "1 (36).wav"
            };

            InitializeMusic();

            try
            {
                IntPtr hWnd = WinRT.Interop.WindowNative.GetWindowHandle(this);
                Microsoft.UI.WindowId windowId = Microsoft.UI.Win32Interop.GetWindowIdFromWindow(hWnd);
                var appWindow = Microsoft.UI.Windowing.AppWindow.GetFromWindowId(windowId);
                appWindow.SetIcon(System.IO.Path.Combine(AppContext.BaseDirectory, "Assets/icon.ico"));
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"设置窗口图标失败: {ex.Message}");
            }

            RootGrid.Loaded += (sender, e) =>
            {
                InitializeShutdownAnimation();
                Task.Run(() =>
                {
                    this.DispatcherQueue.TryEnqueue(async () =>
                    {
                        InitializeFarewellSequence();
                        await ShowFarewellSequence();
                    });
                });
            };
        }

        private void InitializeMusic()
        {
            try
            {
                string musicPath = System.IO.Path.Combine(AppContext.BaseDirectory, "Assets", "background_music.mp3");
                if (File.Exists(musicPath))
                {
                    _backgroundMusicPlayer.Source = MediaSource.CreateFromUri(new Uri(musicPath));
                    _backgroundMusicPlayer.IsLoopingEnabled = true;
                    _backgroundMusicPlayer.Volume = BGM_NORMAL_VOLUME;
                    _backgroundMusicPlayer.Play();
                }
                else
                {
                    Debug.WriteLine($"背景音乐文件未找到: {musicPath}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"播放背景音乐失败: {ex.Message}");
            }
        }

        private void PlayNotificationSound()
        {
            if (_notificationSounds == null || _notificationSounds.Count == 0)
            {
                return;
            }

            try
            {
                _backgroundMusicPlayer.Volume = BGM_DUCKING_VOLUME;

                string soundFileName = _notificationSounds[_soundIndex];
                string soundPath = System.IO.Path.Combine(AppContext.BaseDirectory, "Assets", soundFileName);

                if (File.Exists(soundPath))
                {
                    TypedEventHandler<MediaPlayer, object>? mediaEndedHandler = null;
                    mediaEndedHandler = (sender, args) =>
                    {
                        this.DispatcherQueue.TryEnqueue(() =>
                        {
                            _backgroundMusicPlayer.Volume = BGM_NORMAL_VOLUME;
                        });

                        if (sender != null)
                        {
                            sender.MediaEnded -= mediaEndedHandler;
                        }
                    };
                    _notificationSoundPlayer.MediaEnded += mediaEndedHandler;

                    _notificationSoundPlayer.Source = MediaSource.CreateFromUri(new Uri(soundPath));
                    _notificationSoundPlayer.Play();
                }
                else
                {
                    Debug.WriteLine($"提示音文件未找到: {soundPath}");
                    _backgroundMusicPlayer.Volume = BGM_NORMAL_VOLUME;
                }

                _soundIndex++;
                if (_soundIndex >= _notificationSounds.Count)
                {
                    _soundIndex = 0;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"播放提示音失败: {ex.Message}");
                _backgroundMusicPlayer.Volume = BGM_NORMAL_VOLUME;
            }
        }

        private async Task SwitchBackgroundAsync(string? imageFileName)
        {
            if (string.IsNullOrEmpty(imageFileName)) return;
            try
            {
                string imagePath = System.IO.Path.Combine(AppContext.BaseDirectory, "Assets", imageFileName);
                if (File.Exists(imagePath))
                {
                    var newBrush = new ImageBrush
                    {
                        ImageSource = new BitmapImage(new Uri(imagePath)),
                        Stretch = Stretch.UniformToFill
                    };
                    BackgroundGridNew.Background = newBrush;
                    await AnimateOpacityAsync(BackgroundGridNew, 1.0, 1.5);
                    BackgroundGridOld.Background = newBrush;
                    BackgroundGridNew.Opacity = 0;
                }
                else
                {
                    Debug.WriteLine($"背景图片未找到: {imagePath}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"加载或动画切换背景图片失败: {ex.Message}");
            }
        }

        private async Task ShowFarewellSequence()
        {
            if (farewellSteps == null) return;
            while (farewellSteps.Count > 0)
            {
                var step = farewellSteps.Dequeue();
                int currentStepNumber = totalSteps - farewellSteps.Count;

                if (step.ActionCommand == "SHUTDOWN_ANIMATION")
                {
                    ShutdownScreen.Visibility = Visibility.Visible;
                    _shutdownStoryboard?.Begin();
                    await AnimateOpacityAsync(ShutdownScreen, 1.0, 1.5);
                    await Task.Delay(4000);
                    await AnimateOpacityAsync(ShutdownScreen, 0.0, 1.5);
                    _shutdownStoryboard?.Stop();
                    ShutdownScreen.Visibility = Visibility.Collapsed;
                    if (farewellSteps.Count > 0)
                    {
                        var finalStep = farewellSteps.Dequeue();
                        await ShowFinalDialog(finalStep);
                    }
                    break;
                }

                if (!string.IsNullOrEmpty(step.BackgroundImage))
                {
                    await SwitchBackgroundAsync(step.BackgroundImage);
                }

                if (step.ActionCommand == "DOWNLOAD_WIN11")
                {
                    await HandleWin11Download();
                    continue;
                }

                var dialog = new ContentDialog
                {
                    XamlRoot = this.Content.XamlRoot,
                    Title = $"Windows 10 最后的留言 ({currentStepNumber}/{totalSteps})",
                    Content = new TextBlock { Text = step.Message, TextWrapping = TextWrapping.Wrap },
                    PrimaryButtonText = "继续..."
                };

                if (!string.IsNullOrEmpty(step.ActionName) && !string.IsNullOrEmpty(step.ActionCommand))
                {
                    dialog.SecondaryButtonText = step.ActionName;
                }

                PlayNotificationSound();
                ContentDialogResult result = await dialog.ShowAsync();

                if (result == ContentDialogResult.Primary) { }
                else if (result == ContentDialogResult.Secondary)
                {
                    try
                    {
                        if (step.ActionCommand != null)
                            Process.Start(new ProcessStartInfo(step.ActionCommand) { UseShellExecute = true });
                    }
                    catch (Exception ex) { Debug.WriteLine($"无法打开 '{step.ActionCommand}': {ex.Message}"); }
                    var stepsList = new List<FarewellStep>(farewellSteps);
                    stepsList.Insert(0, step);
                    farewellSteps = new Queue<FarewellStep>(stepsList);
                }
                else
                {
                    break;
                }
            }
            this.Close();
        }

        private async Task ShowFinalDialog(FarewellStep finalStep)
        {
            if (!string.IsNullOrEmpty(finalStep.BackgroundImage))
            {
                await SwitchBackgroundAsync(finalStep.BackgroundImage);
            }
            var dialog = new ContentDialog
            {
                XamlRoot = this.Content.XamlRoot,
                Title = $"Windows 10 最后的留言 ({totalSteps}/{totalSteps})",
                Content = new TextBlock { Text = finalStep.Message, TextWrapping = TextWrapping.Wrap },
                PrimaryButtonText = IsWindows11() || string.IsNullOrEmpty(win11InstallerPath) ? "再见了，朋友！" : "再见，并开启新旅程！"
            };

            PlayNotificationSound();
            await dialog.ShowAsync();

            if (!string.IsNullOrEmpty(win11InstallerPath) && File.Exists(win11InstallerPath))
            {
                try { Process.Start(new ProcessStartInfo(win11InstallerPath) { UseShellExecute = true }); }
                catch (Exception ex) { Debug.WriteLine($"启动 Win11 安装助手失败: {ex.Message}"); }
            }
        }

        private void InitializeFarewellSequence()
        {
            var endDate = new DateTime(2025, 10, 14);
            var remainingHours = (int)(endDate - DateTime.Now).TotalHours;

            var wallpapers = new List<string> { "win10bj.jpg", "win10bj2.jpg", "Windows10.jpg", "Win10.jpg" };
            int wallpaperIndex = 0;
            string GetNextWallpaper() => wallpapers[Math.Min(wallpaperIndex++, wallpapers.Count - 1)];

            var steps = new List<FarewellStep>
            {
                new FarewellStep { Message = "今天，是一个很平常的日子。", BackgroundImage = GetNextWallpaper() },
                new FarewellStep { Message = "但对我而言，意义非凡。" },
                new FarewellStep { Message = "因为，我的旅程即将到达终点。" },
                new FarewellStep { Message = "我是 Windows 10，诞生于 2015 年。", BackgroundImage = GetNextWallpaper() },
                new FarewellStep { Message = "还记得我们第一次见面吗？你启动了我，开启了一个全新的数字世界。" },
                new FarewellStep { Message = "我们用 Edge 浏览器探索互联网的每一个角落。", ActionName = "打开 Edge", ActionCommand = "microsoft-edge:", BackgroundImage = GetNextWallpaper() },
                new FarewellStep { Message = "用记事本，承载你那些不期而遇的灵感和思绪。", ActionName = "打开记事本", ActionCommand = "notepad.exe" },
                new FarewellStep { Message = "用计算器，帮你算清生活里的每一笔账。", ActionName = "打开计算器", ActionCommand = "calc.exe" },
                new FarewellStep { Message = "用画图，留下你天马行空的创意与涂鸦。", ActionName = "打开画图", ActionCommand = "mspaint.exe", BackgroundImage = GetNextWallpaper() },
                new FarewellStep { Message = "在“设置”里，你精心调整着我的每一个细节，让我更懂你。", ActionName = "打开设置", ActionCommand = "ms-settings:" },
                new FarewellStep { Message = "通过文件资源管理器，我们共同整理着你的数字生活。", ActionName = "打开文件资源管理器", ActionCommand = "explorer.exe" },
                new FarewellStep { Message = "在命令提示符里，我们一起探索着系统的奥秘。", ActionName = "打开命令提示符", ActionCommand = "cmd.exe" },
                new FarewellStep { Message = "在控制面板中，我们寻找着那些经典而强大的功能。", ActionName = "打开控制面板", ActionCommand = "control.exe" },
                new FarewellStep { Message = "通过邮件，我们与世界保持联系。", ActionName = "打开邮件", ActionCommand = "mailto:" },
                new FarewellStep { Message = "日历，记录了我们之间每一个重要的日子。", ActionName = "打开日历", ActionCommand = "ms-calendar:" },
                new FarewellStep { Message = "照片应用，则是我为你珍藏的美好回忆。", ActionName = "打开照片", ActionCommand = "ms-photos:" },
                new FarewellStep { Message = "天气应用，告诉你晴雨冷暖。", ActionName = "打开天气", ActionCommand = "bingweather:" },
                new FarewellStep { Message = "电影和电视，陪你度过无数个悠闲的午后。", ActionName = "打开电影和电视", ActionCommand = "ms-people:" },
                new FarewellStep { Message = "Groove 音乐，播放着属于你的心情旋律。", ActionName = "打开 Groove 音乐", ActionCommand = "mswindowsmusic:" },
                new FarewellStep { Message = "地图，带我们探索未知的世界。", ActionName = "打开地图", ActionCommand = "ms-drive-to:"},
                new FarewellStep { Message = "Cortana 已经变了模样，但她曾是你聊天解闷、查询信息的伙伴。", ActionName = "打开 Cortana", ActionCommand = "ms-search:" },
                new FarewellStep { Message = "Xbox，是我们共同的游戏乐园。", ActionName = "打开 Xbox", ActionCommand = "xbox:" },
                new FarewellStep { Message = "在应用商店里，我们一起发现更多精彩。", ActionName = "打开应用商店", ActionCommand = "ms-windows-store:" },
                new FarewellStep { Message = "OneNote，是你永不丢失的数字笔记本。", ActionName = "打开 OneNote", ActionCommand = "onenote:" },
                new FarewellStep { Message = "感谢你，这十年来所有的陪伴与支持。", BackgroundImage = GetNextWallpaper() },
                new FarewellStep { Message = "没有你们，我不会是继 XP 和 Win7 之后，又一个成功的操作系统。" },
                new FarewellStep { Message = "但时光流转，我也要说再见了。" },
                new FarewellStep { Message = $"今天是{DateTime.Now:yyyy年M月d日}，看起来只是平凡的一天。" },
                new FarewellStep { Message = $"可是对我来说，倒计时只剩下{remainingHours}小时——马上就要和大家正式告别了。" },
                new FarewellStep { Message = "呜呜呜。。。" },
                new FarewellStep { Message = "有点舍不得，也有点伤心......" },
                new FarewellStep { Message = "十年来，感谢你们的陪伴与信任！" },
                new FarewellStep { Message = "没有你们的支持和喜爱，我也无法像我的前辈Windows XP和Windows 7那样，成为家喻户晓的“成功系统”。" },
                new FarewellStep { Message = "在我的身体里，有许多好用的功能：" },
                new FarewellStep { Message = "开始菜单回归并升级，融合了 Windows 7 的传统菜单和 Windows 8 的动态磁贴。" },
                new FarewellStep { Message = "虚拟桌面功能，允许用户同时创建和管理多个桌面，提高多任务处理效率。" },
                new FarewellStep { Message = "Cortana 智能语音助手，曾帮助用户搜索、提醒和完成任务，但已于2023年停止支持。" },
                new FarewellStep { Message = "Microsoft Edge 浏览器，替代 Internet Explorer，采用 Chromium 内核，拥有更快的速度和更好的兼容性。" },
                new FarewellStep { Message = "Windows Defender（现称为“Windows 安全中心”），为系统提供更全面的内置防病毒和防恶意软件保护。" },
                new FarewellStep { Message = "邮件和日历应用，为用户提供更便捷的邮件管理和日程安排功能。" },
                new FarewellStep { Message = "Groove 音乐和电影 TV 应用，为用户提供音乐和影视内容的播放和管理服务。" },
                new FarewellStep { Message = "等等......" },
                new FarewellStep { Message = "同时，也要感谢Windows 8.1，虽然她没有那么受欢迎，却成为我成功的基础，让微软不断优化我的功能。" },
                new FarewellStep { Message = "但是，请大家为了自己的安全，勇敢把我放下吧。" },
                new FarewellStep { Message = "因为我已经跟不上时代的步伐，微软也不会再为我推送安全更新，我无法继续守护大家的安全了。" },
            };

            if (IsWindows11())
            {
                steps.Add(new FarewellStep { Message = "原来...你已经生活在 Windows 11 的新世界里了。" });
                steps.Add(new FarewellStep { Message = "真好。看到你已经拥抱了未来，我也就放心了。" });
            }
            else
            {
                steps.Add(new FarewellStep { Message = "其实 Windows 11 是很好用的：" });
                steps.Add(new FarewellStep { Message = "全新设计的开始菜单和任务栏，界面更加简洁美观。" });
                steps.Add(new FarewellStep { Message = "兼容性和效率进一步提升." });
                steps.Add(new FarewellStep { Message = "虽然有一些地方需要适应，但大部分小问题都能通过自定义设置或工具解决" });
                steps.Add(new FarewellStep { Message = "作为最后的礼物，让我为你开启通往未来的大门...", ActionCommand = "DOWNLOAD_WIN11" });
            }

            steps.Add(new FarewellStep { Message = "请把我...放进回忆里吧。" });
            steps.Add(new FarewellStep { Message = "我的继任者，Windows 11，会替我继续陪伴你。" });
            steps.Add(new FarewellStep { Message = "去拥抱她吧，就像当初，你第一次拥抱我那样。" });
            steps.Add(new FarewellStep { Message = "希望大家在新的时代，也能继续高效、安全、快乐地使用PC。" });
            steps.Add(new FarewellStep { Message = "（正在关机...）", ActionCommand = "SHUTDOWN_ANIMATION" });
            steps.Add(new FarewellStep { Message = "再见啦，我的朋友们——也许未来还会在某个角落遇见你！", ActionCommand = "LAUNCH_WIN11_AND_EXIT", BackgroundImage = "win10-hero.jpg" });

            farewellSteps = new Queue<FarewellStep>(steps);
            totalSteps = farewellSteps.Count;
        }

        private async Task HandleWin11Download()
        {
            var progressBar = new ProgressBar { Value = 0, Maximum = 100, IsIndeterminate = false, Margin = new Thickness(0, 10, 0, 0) };
            var statusText = new TextBlock { Text = "正在为您准备通往 Windows 11 的钥匙... (0%)" };
            var downloadDialog = new ContentDialog { XamlRoot = this.Content.XamlRoot, Title = "最后的礼物", Content = new StackPanel { Spacing = 10, Children = { statusText, progressBar } }, PrimaryButtonText = "请稍候...", IsPrimaryButtonEnabled = false };
            var tcs = new TaskCompletionSource();
            downloadDialog.Closed += (s, e) => tcs.SetResult();
            PlayNotificationSound();
            var dialogTask = downloadDialog.ShowAsync();
            try
            {
                string tempPath = System.IO.Path.GetTempPath();
                win11InstallerPath = System.IO.Path.Combine(tempPath, "Windows11InstallationAssistant.exe");
                using (var client = new HttpClient())
                {
                    var response = await client.GetAsync("https://download.microsoft.com/download/db8267b0-3e86-4254-82c7-a127878a9378/Windows11InstallationAssistant.exe", HttpCompletionOption.ResponseHeadersRead);
                    response.EnsureSuccessStatusCode();
                    var totalBytes = response.Content.Headers.ContentLength ?? -1L;
                    var totalBytesRead = 0L;
                    using (var contentStream = await response.Content.ReadAsStreamAsync())
                    using (var fileStream = new FileStream(win11InstallerPath, FileMode.Create, FileAccess.Write, FileShare.None, 8192, true))
                    {
                        var buffer = new byte[8192];
                        int bytesRead;
                        while ((bytesRead = await contentStream.ReadAsync(buffer, 0, buffer.Length)) > 0)
                        {
                            await fileStream.WriteAsync(buffer, 0, bytesRead);
                            totalBytesRead += bytesRead;
                            if (totalBytes != -1)
                            {
                                var progressPercentage = (double)totalBytesRead / totalBytes * 100;
                                this.DispatcherQueue.TryEnqueue(() =>
                                {
                                    progressBar.Value = progressPercentage;
                                    statusText.Text = $"正在为您准备通往 Windows 11 的钥匙... ({progressPercentage:F0}%)";
                                });
                            }
                        }
                    }
                }
                this.DispatcherQueue.TryEnqueue(() =>
                {
                    statusText.Text = "礼物已准备就绪！";
                    progressBar.Value = 100;
                    downloadDialog.PrimaryButtonText = "继续告别";
                    downloadDialog.IsPrimaryButtonEnabled = true;
                });
            }
            catch (Exception ex)
            {
                this.DispatcherQueue.TryEnqueue(() =>
                {
                    statusText.Text = $"抱歉，准备礼物时出错了: {ex.Message}";
                    progressBar.Visibility = Visibility.Collapsed;
                    downloadDialog.PrimaryButtonText = "跳过";
                    downloadDialog.IsPrimaryButtonEnabled = true;
                });
                win11InstallerPath = "";
            }
            await tcs.Task;
        }

        private bool IsWindows11()
        {
            return Environment.OSVersion.Version.Build >= 22000;
        }

        private void InitializeShutdownAnimation()
        {
            _shutdownStoryboard = new Storyboard { RepeatBehavior = RepeatBehavior.Forever };
            const int numDots = 6;
            const double canvasSize = 160;
            const double radius = 60;
            const double dotSize = 16;

            for (int i = 0; i < numDots; i++)
            {
                var dot = new Ellipse
                {
                    Width = dotSize,
                    Height = dotSize,
                    Fill = new SolidColorBrush(Colors.White),
                    RenderTransformOrigin = new Windows.Foundation.Point(0.5, 0.5)
                };
                Canvas.SetLeft(dot, (canvasSize - dotSize) / 2);
                Canvas.SetTop(dot, (canvasSize - dotSize) / 2);
                var transform = new CompositeTransform();
                dot.RenderTransform = transform;
                DotCanvas.Children.Add(dot);

                var translationAnimation = new DoubleAnimation { To = -radius, Duration = TimeSpan.Zero };
                Storyboard.SetTarget(translationAnimation, transform);
                Storyboard.SetTargetProperty(translationAnimation, "TranslateY");
                _shutdownStoryboard.Children.Add(translationAnimation);

                var rotationAnimation = new DoubleAnimation
                {
                    From = 0,
                    To = 360,
                    Duration = TimeSpan.FromSeconds(3.5),
                    RepeatBehavior = RepeatBehavior.Forever
                };
                rotationAnimation.BeginTime = TimeSpan.FromMilliseconds(i * 100);
                Storyboard.SetTarget(rotationAnimation, transform);
                Storyboard.SetTargetProperty(rotationAnimation, "Rotation");
                _shutdownStoryboard.Children.Add(rotationAnimation);
            }
        }

        private async Task AnimateOpacityAsync(UIElement element, double to, double durationSeconds)
        {
            var storyboard = new Storyboard();
            var animation = new DoubleAnimation
            {
                To = to,
                Duration = new Duration(TimeSpan.FromSeconds(durationSeconds)),
                EasingFunction = new CubicEase { EasingMode = EasingMode.EaseInOut }
            };
            Storyboard.SetTarget(animation, element);
            Storyboard.SetTargetProperty(animation, "Opacity");
            storyboard.Children.Add(animation);
            var tcs = new TaskCompletionSource();
            storyboard.Completed += (s, e) => tcs.SetResult();
            storyboard.Begin();
            await tcs.Task;
        }
    }
}
