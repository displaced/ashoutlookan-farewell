using LibVLCSharp.Shared;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using YoutubeExtractor;

namespace AshoutlookanFarewell
{
    /// <summary>
    /// Controls downloading of media, playback, volume alterations, etc.
    /// </summary>
    public static class MusicController
    {
        static LibVLC libvlc;
        static MediaPlayer player;
        

        /// <summary>
        /// Downloads media from YouTube.
        /// </summary>
        /// <param name="url">The YT URL of the media to get.</param>
        /// <param name="destinationFile">The local file to save to</param>
        public static void AcquireTrack(string url, string destinationFile)
        {
            IEnumerable<VideoInfo> videoInfos = DownloadUrlResolver.GetDownloadUrls(url);

            VideoInfo video = videoInfos
                    .OrderByDescending(info => info.AudioBitrate)
                    .First();

            if (video.RequiresDecryption)
            {
                DownloadUrlResolver.DecryptDownloadUrl(video);
            }

            var videoDownloader = new VideoDownloader(video, destinationFile);
            videoDownloader.Execute();
        }


        /// <summary>
        /// Get the local path to a previously-downloaded media file.  If the file isn't found, AcquireTrack() is called to attempt
        /// to download it.
        /// </summary>
        public static string GetMediaFile()
        {
            var expectedFile = Path.Combine(Settings.DataDir, "ashoutlookan-farewell");
            if (!File.Exists(expectedFile))
            {
                AcquireTrack(Settings.MediaUrl, expectedFile);
            }

            return expectedFile;
        }


        /// <summary>
        /// Initialise the MusicController
        /// </summary>
        public static void Init()
        {
            // TODO: Figure out how to find the path to libvlc when deployed as a VSTO extension
            //       The default 'publish' feature isn't properly bundling these dependencies.
            Core.Initialize(@"C:\Users\chris.platts\source\repos\AshoutlookanFarewell\AshoutlookanFarewell\bin\Debug\libvlc\win-x64");
            libvlc = new LibVLC(enableDebugLogs: true);

            var file = GetMediaFile();
            var media = new Media(libvlc, new Uri(file));
            player = new MediaPlayer(media);
        }


        /// <summary>
        /// Begin playback.  Should be called in response to composing a new missive, or reading a received one. 
        /// Only works with emails, does not include carrier-pidgeon detection logic.
        /// </summary>
        public static void StartPlayback()
        {
            player.Volume = 100;   
            
            var success = player.Play();
            if (!success)
            {
                // Should do something useful if there's an error.
                var err = libvlc.LastLibVLCError;
            }
        }


        /// <summary>
        /// Stops playback.  Should be called when closing a missive.
        /// </summary>
        public static void StopPlayback()
        {
            if (player != null)
            {
                // Fade out.  This is nasty and will block the UI whilst fading.
                while (player.Volume > 0)
                {
                    player.Volume -= 2;
                    Thread.Sleep(25);
                }

                player.Stop();
            }
        }

    }
}
