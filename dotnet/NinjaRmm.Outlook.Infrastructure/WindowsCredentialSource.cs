using System;
using System.Runtime.InteropServices;

namespace NinjaRmm.Outlook.Infrastructure
{
    // Reads an existing generic entry only. Never creates, overwrites or deletes credentials.
    public sealed class WindowsCredentialSource : ICredentialSource
    {
        private readonly string target;
        public WindowsCredentialSource(string target)
        {
            if (string.IsNullOrWhiteSpace(target)) throw new ArgumentException("Credential target is missing.");
            this.target = target;
        }
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct Credential
        {
            public uint Flags, Type;
            public string TargetName, Comment;
            public System.Runtime.InteropServices.ComTypes.FILETIME LastWritten;
            public uint CredentialBlobSize;
            public IntPtr CredentialBlob;
            public uint Persist, AttributeCount;
            public IntPtr Attributes;
            public string TargetAlias, UserName;
        }
        [DllImport("advapi32.dll", EntryPoint = "CredReadW", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool CredRead(string target, uint type, uint flags, out IntPtr credential);
        [DllImport("advapi32.dll")] private static extern void CredFree(IntPtr credential);
        public ClientCredential Read()
        {
            IntPtr pointer;
            if (!CredRead(target, 1, 0, out pointer)) throw new InvalidOperationException("Credential Manager entry is unavailable.");
            try
            {
                var credential = (Credential)Marshal.PtrToStructure(pointer, typeof(Credential));
                if (credential.CredentialBlobSize == 0 || credential.CredentialBlobSize % 2 != 0 || credential.CredentialBlobSize > 5120)
                    throw new InvalidOperationException("Credential Manager entry has an invalid secret format.");
                return new ClientCredential(credential.UserName,
                    Marshal.PtrToStringUni(credential.CredentialBlob, checked((int)credential.CredentialBlobSize / 2)));
            }
            finally { CredFree(pointer); }
        }
    }
}
