namespace NinjaRmm.Outlook.Core
{
    public enum MatchRule { None, NinjaId, ExternalNumber, Conversation, Subject, Learned }
    public enum MatchLevel { Off, LogOnly, Confirm, Automatic }

    // Instance-owned configuration avoids the process-global VBA test overrides.
    public sealed class MatchLevels
    {
        private readonly int ninjaId, externalNumber, conversation, subject, learned, cap;

        public MatchLevels(int ninjaId = 1, int externalNumber = 1, int conversation = 1,
            int subject = 1, int learned = 1, int cap = 3)
        {
            this.ninjaId = ninjaId;
            this.externalNumber = externalNumber;
            this.conversation = conversation;
            this.subject = subject;
            this.learned = learned;
            this.cap = cap;
        }

        public MatchLevel For(MatchRule rule)
        {
            int configured;
            switch (rule)
            {
                case MatchRule.NinjaId: configured = ninjaId; break;
                case MatchRule.ExternalNumber: configured = externalNumber; break;
                case MatchRule.Conversation: configured = conversation; break;
                case MatchRule.Subject: configured = subject; break;
                case MatchRule.Learned: configured = learned; break;
                default: return MatchLevel.Off;
            }
            return (MatchLevel)System.Math.Max(0, System.Math.Min(3, System.Math.Min(configured, cap)));
        }

        public bool AnyAt(MatchLevel level)
        {
            return For(MatchRule.NinjaId) == level || For(MatchRule.ExternalNumber) == level ||
                For(MatchRule.Conversation) == level || For(MatchRule.Subject) == level ||
                For(MatchRule.Learned) == level;
        }

        public bool AnyActive
        {
            get
            {
                return For(MatchRule.NinjaId) != MatchLevel.Off ||
                    For(MatchRule.ExternalNumber) != MatchLevel.Off ||
                    For(MatchRule.Conversation) != MatchLevel.Off ||
                    For(MatchRule.Subject) != MatchLevel.Off ||
                    For(MatchRule.Learned) != MatchLevel.Off;
            }
        }
    }
}
