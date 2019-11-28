using Microsoft.SharePoint.Client;
using NLog;
using OfficeDevPnP.Core;
using System;
using System.Diagnostics;

namespace TokenHelper
{
    public partial class TokenHelper : MetroFramework.Forms.MetroForm
    {
        Logger log = LogManager.GetLogger(typeof(Program).FullName);

        public TokenHelper()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            log.Info("Hello world, the form has been loaded");

            HideAdfsParameterControls();
            HideHighTrustParameterControls();
        }

        private void ShowAdfsParameterControls()
        {
            lbSpDomain.Show();
            txtSpDomain.Show();
            lbSpSts.Show();
            txtSpSts.Show();
            lbSpIdp.Show();
            txtSpIdp.Show();
            lbSpTokenExpirationWindow.Show();
            txtSpTokenExpirationWindow.Show();
        }

        private void HideAdfsParameterControls()
        {
            lbSpDomain.Hide();
            txtSpDomain.Hide();
            lbSpSts.Hide();
            txtSpSts.Hide();
            lbSpIdp.Hide();
            txtSpIdp.Hide();
            lbSpTokenExpirationWindow.Hide();
            txtSpTokenExpirationWindow.Hide();
        }

        private void ShowHighTrustParameterControls()
        {
            lbSpCertificateIssuerId.Show();
            txtSpCertificateIssuerId.Show();
            lbSpCertificatePath.Show();
            txtSpCertificatePath.Show();
            lbSpCertificatePasswords.Show();
            txtSpCertificatePasswords.Show();
            lbSpAppSecret.Hide();
            txtSpAppSecret.Hide();
        }

        private void HideHighTrustParameterControls()
        {
            lbSpCertificateIssuerId.Hide();
            txtSpCertificateIssuerId.Hide();
            lbSpCertificatePath.Hide();
            txtSpCertificatePath.Hide();
            lbSpCertificatePasswords.Hide();
            txtSpCertificatePasswords.Hide();
            lbSpAppSecret.Show();
            txtSpAppSecret.Show();
        }

        private void BtnGo_Click(object sender, EventArgs e)
        {
            try
            {

                //
                // Begin - SharePoint Online 
                //

                if (toggleSpoCredentials.Checked)
                {
                    // SharePoint Online Credentials 
                    log.Info("Getting the client context now using GetSharePointOnlineAuthenticatedContextTenant sharepoint online credentials");
                    using (ClientContext cc = new AuthenticationManager().GetSharePointOnlineAuthenticatedContextTenant(txtSpoSiteUrlCredentials.Text, txtSpoUsername.Text, txtSpoPasswords.Text))
                    {
                        Web web = cc.Web;
                        cc.Load(web, w => w.Title);
                        cc.ExecuteQueryRetry();

                        lbHint.Text = System.String.Format("You've just got the token, the site title is {0}", web.Title);
                    }
                }

                if (toggleSpoAppOnly.Checked)
                {
                    // SharePoint Online App Only 
                    log.Info("Getting the client context now using GetAppOnlyAuthenticatedContext sharepoint online app only");
                    using (ClientContext cc = new AuthenticationManager().GetAppOnlyAuthenticatedContext(txtSpoSiteUrlAppOnly.Text, txtSpoAppId.Text, txtSpoAppSecret.Text))
                    {
                        Web web = cc.Web;
                        cc.Load(web, w => w.Title);
                        cc.ExecuteQueryRetry();

                        lbHint.Text = System.String.Format("You've just got the token, the site title is {0}", web.Title);
                    }
                }

                if (toggleSpoInteractive.Checked)
                {
                    // SharePoint Online Interactive 
                    log.Info("Getting the client context now using GetWebLoginClientContext sharepoint online interactive");
                    using (ClientContext cc = new AuthenticationManager().GetWebLoginClientContext(txtSpoSiteUrlInteractive.Text))
                    {
                        Web web = cc.Web;
                        cc.Load(web, w => w.Title);
                        cc.ExecuteQueryRetry();

                        lbHint.Text = System.String.Format("You've just got the token, the site title is {0}", web.Title);
                    }
                }

                //
                // End - SharePoint Online 
                //

                //
                // Begin - SharePoint On Prem 
                //

                if (toggleSpCredentials.Checked && !checkADFS.Checked)
                {
                    // SharePoint On Prem Credentials, without ADFS 
                    log.Info("Getting the client context now using GetNetworkCredentialAuthenticatedContext  sharepoint on prem credentials, without ADFS");
                    using (ClientContext cc = new AuthenticationManager().GetNetworkCredentialAuthenticatedContext(txtSpSiteUrlCredentials.Text, txtSpUsername.Text, txtSpPasswords.Text, txtSpDomain.Text))
                    {
                        Web web = cc.Web;
                        cc.Load(web, w => w.Title);
                        cc.ExecuteQueryRetry();

                        lbHint.Text = System.String.Format("You've just got the token, the site title is {0}", web.Title);
                    }
                }

                if (toggleSpCredentials.Checked && checkADFS.Checked)
                {
                    // SharePoint On Prem Credentials, with ADFS 
                    log.Info("Getting the client context now using GetADFSUserNameMixedAuthenticatedContext  sharepoint on prem credentials, with ADFS");
                    using (ClientContext cc = new AuthenticationManager().GetADFSUserNameMixedAuthenticatedContext(txtSpSiteUrlCredentials.Text, txtSpUsername.Text, txtSpPasswords.Text, txtSpDomain.Text, txtSpSts.Text, lbSpIdp.Text, lbSpTokenExpirationWindow.Text.ToInt32()))
                    {
                        Web web = cc.Web;
                        cc.Load(web, w => w.Title);
                        cc.ExecuteQueryRetry();

                        lbHint.Text = System.String.Format("You've just got the token, the site title is {0}", web.Title);
                    }
                }


                if (toggleSpAppOnly.Checked && checkHighTrust.Checked)
                {
                    // SharePoint On Prem App Only, High Trust 
                    log.Info("Getting the client context now using GetHighTrustCertificateAppOnlyAuthenticatedContext  sharepoint on prem credentials, high trust");
                    using (ClientContext cc = new AuthenticationManager().GetHighTrustCertificateAppOnlyAuthenticatedContext(txtSpSiteUrlAppOnly.Text, txtSpAppId.Text, txtSpCertificatePath.Text, txtSpCertificatePasswords.Text, txtSpCertificateIssuerId.Text))
                    {
                        Web web = cc.Web;
                        cc.Load(web, w => w.Title);
                        cc.ExecuteQueryRetry();

                        lbHint.Text = System.String.Format("You've just got the token, the site title is {0}", web.Title);
                    }
                }

                if (toggleSpAppOnly.Checked && !checkHighTrust.Checked)
                {
                    // SharePoint On Prem App Only, Low Trust 
                    log.Info("Getting the client context now using GetAppOnlyAuthenticatedContext  sharepoint on prem credentials, low trust");
                    using (ClientContext cc = new AuthenticationManager().GetAppOnlyAuthenticatedContext(txtSpSiteUrlAppOnly.Text, txtSpAppId.Text, txtSpAppSecret.Text))
                    {
                        Web web = cc.Web;
                        cc.Load(web, w => w.Title);
                        cc.ExecuteQueryRetry();

                        lbHint.Text = System.String.Format("You've just got the token, the site title is {0}", web.Title);
                    }
                }

                //
                // End - SharePoint On Prem 
                //

                //
                // Begin - Azure Ad  
                //


                // ... 


                //
                // End - Azure Ad  
                //
            }
            catch (Exception ex)
            {
                lbHint.Text = "Something went wrong, pls check the log file for more details";
                log.Error(ex, "Getting errors while BtnGo_Click");
            }
        }

        private void BtnLogs_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@".\Logs");
                lbHint.Text = "";
            }
            catch (Exception ex)
            {
                log.Error(ex, "Getting errors while BtnLogs_Click");
                lbHint.Text = "Something went wrong, pls check the config file";
            }
        }

        private void CheckADFS_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (!checkADFS.Checked)
                {
                    HideAdfsParameterControls();
                }

                if (checkADFS.Checked)
                {
                    ShowAdfsParameterControls();
                }
            }
            catch (Exception ex)
            {
                log.Error(ex, "Getting errors while CheckADFS_CheckedChanged");
            }
        }

        private void CheckHighTrust_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (checkHighTrust.Checked)
                {
                    ShowHighTrustParameterControls();
                }

                if (!checkHighTrust.Checked)
                {
                    HideHighTrustParameterControls();
                }
            }
            catch (Exception ex)
            {
                log.Error(ex, "Getting errors while CheckHighTrust_CheckedChanged");
            }
        }
    }
}
