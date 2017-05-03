﻿//------------------------------------------------------------------------------
// <copyright file="Command.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace VSPluginSample
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class Command
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("bab231d0-d39b-489b-920c-98e6b23e005a");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly CommandPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="Command"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private Command(CommandPackage package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
                commandService.AddCommand(menuItem);
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static Command Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(CommandPackage package)
        {
            Instance = new Command(package);
        }

        /// <summary>
        /// Print to Output Window
        /// </summary>
        internal void OutputString(string output)
        {
            var outPutPane = this.package.OutputPane;
            outPutPane.OutputString(output + Environment.NewLine);
        }

        /// <summary>
        /// Print separater
        /// </summary>
        internal void PrintSeparater()
        {
            this.OutputString(new string('-', 80) + Environment.NewLine);
        }

        /// <summary>
        /// Print active document
        /// </summary>
        internal void PrintActiveDocument()
        {
            var dte = this.package.GetDTE();
            try
            {
                var doument = dte.ActiveDocument;
                if (doument != null)
                {
                    var filename = doument.FullName;
                    this.OutputString(filename + Environment.NewLine);
                }
                else
                {
                    this.OutputString("There is no active document" + Environment.NewLine);
                }
            }
            catch (ArgumentException)
            {
                this.OutputString("There is no active document" + Environment.NewLine);
            }
        }

        /// <summary>
        /// Print Projects
        /// </summary>
        internal void PrintProjects()
        {
            var dte = this.package.GetDTE();

            // enumerate Projects 
            foreach (EnvDTE.Project project in dte.Solution)
            {
                this.OutputString(project.Name + ": " + project.FullName + Environment.NewLine);
            }
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            // activate Output Window
            var outPutPane = this.package.OutputPane;
            outPutPane.Activate();

            // clear Output Window
            outPutPane.Clear();

            // print separater
            PrintSeparater();

            // print about Active Document
            PrintActiveDocument();

            // print separater
            PrintSeparater();

            // print Projects
            PrintProjects();
        }
    }
}
