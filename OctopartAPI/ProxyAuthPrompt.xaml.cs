//-----------------------------------------------------------------------
// <copyright file="ProxyAuthPrompt.xaml.cs" company="Octopart">
//     Copyright (c) Octopart. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System.Windows;

namespace OctopartApi
{
    /// <summary>
    /// Interaction logic for ProxyAuthPrompt.xaml
    /// </summary>
    public partial class ProxyAuthPrompt : Window
    {
        /// <summary>
        /// Forces the prompt to be threadsafe.
        /// </summary>
        /// <remarks>
        /// Refer to http://stackoverflow.com/questions/2463822/threading-errors-with-application-loadcomponent-key-already-exists for more information.
        /// </remarks>
        public static object ComponentLock = new object();

        /// <summary>
        /// Initializes a new instance of the PrxyAuthPrompt class
        /// </summary>
        /// <param name="server">The proxy server url, to display to the user</param>
        public ProxyAuthPrompt(string server)
        {
            lock (ComponentLock)
            {
                InitializeComponent();
                textMessage.Text = string.Format(textMessage.Text, server);
                textboxUsername.Focus();
            }
        }

        /// <summary>
        /// Gets the user specified username
        /// </summary>
        public string User
        {
            get { return textboxUsername.Text; }
        }

        /// <summary>
        /// Gets the user specified password
        /// </summary>
        public string Pass
        {
            get { return textboxPassword.Password; }
        }

        /// <summary>
        /// Quits the form
        /// </summary>
        /// <param name="sender">Sender</param>
        /// <param name="e">Event Args</param>
        private void buttonCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        /// <summary>
        /// Quits the form
        /// </summary>
        /// <param name="sender">Sender</param>
        /// <param name="e">Event Args</param>
        private void buttonOK_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }
    }
}
