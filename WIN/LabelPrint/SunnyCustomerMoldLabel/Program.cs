// <copyright file="Program.cs" company="Euro Plus">
//     Copyright © Euro Plus 2014.
// </copyright>
// <summary>This is the Program class.</summary>
//-----------------------------------------------------------------------
namespace SDKSimpleApp
{
    using System;
    using System.IO;
    using System.Reflection;
    using System.Windows.Forms;

    /// <summary>
    /// The Program class.
    /// </summary>
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        internal static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
