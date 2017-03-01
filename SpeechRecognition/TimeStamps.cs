using System;

namespace SpeechRecognition
{
    public class TimeSpans
    {
        public TimeSpan StartTime { get; set; }
        public TimeSpan StopTime { get; set; }

        private TimeSpans(TimeSpan startTime, TimeSpan stopTime)
        {
            StartTime = startTime;
            StopTime = stopTime;
        }

        public static TimeSpans Create(TimeSpan startTime, TimeSpan stopTime)
        {
            return new TimeSpans(startTime, stopTime);
        }
    }
}
