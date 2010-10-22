/*
 * The MIT License
 *
 * Copyright (c) 2010, Manufacture Française des Pneumatiques Michelin, Thomas Maurel
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 */

package com.michelin.cio.hudson.plugins.qc.client;

import com.michelin.cio.hudson.plugins.qc.Messages;
import com.michelin.cio.hudson.plugins.qc.QualityCenterUtils;
import hudson.AbortException;
import hudson.Extension;
import hudson.FilePath;
import hudson.Launcher;
import hudson.model.Hudson;
import hudson.model.Node;
import hudson.model.TaskListener;
import hudson.tools.ToolInstallation;
import hudson.tools.ToolInstaller;
import hudson.tools.ToolInstallerDescriptor;
import hudson.util.ArgumentListBuilder;
import hudson.util.FormValidation;
import java.io.FileInputStream;

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;
import java.io.PrintStream;
import java.net.URL;
import org.apache.commons.io.IOUtils;
import org.kohsuke.stapler.DataBoundConstructor;
import org.kohsuke.stapler.QueryParameter;

/**
 * @author Thomas Maurel
 */
public class QualityCenterClientInstaller extends ToolInstaller {

  private final String qcServerURL;
  private final String localPathToQCClient;

  public static final String INSTALLER_PATH_ON_SERVER = "PlugIns/ClientSideInstallation/";
  public static final String DLL_PATH_ON_NODE = "MERCURY INTERACTIVE\\Quality Center";
  public static final String INSTALLLER_FILE_NAME = "QCClient.msi";
  public static final String DLL_NAME = "OTAClient.dll";

  @DataBoundConstructor
  public QualityCenterClientInstaller(String qcServerURL, String localPathToQCClient) {
      super(null);
      this.qcServerURL = qcServerURL;
      this.localPathToQCClient = localPathToQCClient;
  }

  public String getLocalPathToQCClient() {
    return localPathToQCClient;
  }

  public String getQcServerURL() {
    return qcServerURL;
  }

  @Override
  public FilePath performInstallation(ToolInstallation tool, Node node, TaskListener log) throws IOException, InterruptedException {
    FilePath expectedLocation = preferredLocation(tool, node);
    PrintStream out = log.getLogger();
    FilePath qcFolder = expectedLocation.child(DLL_PATH_ON_NODE);

    // Check if it has been installed by Hudson or manually
    // Might be in the QC_HOME, or in QC_HOME\MERCURY INTERACTIVE\Quality Center
    if(expectedLocation.child(".installedByHudson").exists() || expectedLocation.child(DLL_NAME).exists()) {
      return expectedLocation;
    }

    if(qcFolder.child(".installedByHudson").exists() || qcFolder.child(DLL_NAME).exists()) {
      return qcFolder;
    }

    // Create the directory tree
    expectedLocation.mkdirs();
    FilePath file = expectedLocation.child(INSTALLLER_FILE_NAME);

    // If the installer is stored on the master, copy it from there
    if(localPathToQCClient != null & localPathToQCClient.length() > 0) {
      FilePath installerOnMaster = new FilePath(Hudson.MasterComputer.localChannel, localPathToQCClient);
      if(!installerOnMaster.exists()) {
        log.fatalError(Messages.QualityCenterClientInstaller_CannotFindInstaller());
        throw new AbortException();
      }

      if(installerOnMaster.isDirectory()) {
        log.fatalError(Messages.QualityCenterClientInstaller_ShouldBeAFile());
        throw new AbortException();
      }

      out.println(Messages.QualityCenterClientInstaller_CopyingFromMaster(localPathToQCClient));
      file.copyFrom(installerOnMaster);
    }
    // Else, download it from Quality Center Server
    else {
      String url;
      // Add a "/" at the end of the URL if it isnt already
      if(qcServerURL.lastIndexOf("/") == qcServerURL.length() - 1) {
        url = qcServerURL + INSTALLER_PATH_ON_SERVER + INSTALLLER_FILE_NAME;
      }
      else {
        url = qcServerURL + "/" + INSTALLER_PATH_ON_SERVER + INSTALLLER_FILE_NAME;
      }
      URL installURL = new URL(url);
      out.println(Messages.QualityCenterClientInstaller_Downloading(installURL));
      file.copyFrom(installURL);
    }

    // Perform installation
    install(node.createLauncher(log), log, expectedLocation.absolutize().getRemote(), file);

    // Successfully installed
    // Remove the installer
    file.delete();

    // Check if the DLL exists after installation
    if(!qcFolder.child(DLL_NAME).exists()) {
      log.fatalError(Messages.QualityCenterClientInstaller_CouldntFindDllAfterInstall(DLL_NAME));
      throw new AbortException();
    }
    
    qcFolder.child(".installedByHudson").touch(System.currentTimeMillis());
    
    return qcFolder;
  }

  public void install(Launcher launcher, TaskListener log, String expectedLocation, FilePath qcBundle) throws IOException, InterruptedException {
    PrintStream out = log.getLogger();
    String qcBundleName = qcBundle.getName();
    String qcBundlePath = qcBundle.absolutize().getRemote();

    out.println(Messages.QualityCenterClientInstaller_Installing(qcBundleName));
    String logFile = qcBundlePath + ".install.log";

    StringBuilder cmd = new StringBuilder();
    cmd.append('"').append(qcBundlePath).append('"');
    cmd.append(" /qn /norestart TARGETDIR=");
    cmd.append('"').append(expectedLocation).append('"');
    cmd.append(" /l ");
    cmd.append('"').append(logFile).append('"');

    // Run the install on the node
    // The result must be 0
    ArgumentListBuilder args = new ArgumentListBuilder().add("cmd.exe", "/C").addQuoted(cmd.toString());
    if(launcher.launch().cmds(args).stdout(out).pwd(expectedLocation).join()!=0) {
      log.fatalError(Messages.QualityCenterClientInstaller_AbortedInstall());
      // log file is in UTF-16
      InputStream is= new FileInputStream(logFile);
      InputStreamReader in = new InputStreamReader(is, "UTF-16");
      try {
          IOUtils.copy(in,new OutputStreamWriter(out));
      } finally {
          in.close();
          is.close();
      }
      throw new AbortException();
    }

    out.println(Messages.QualityCenterClientInstaller_InstallationSuccessfull());
  }

  @Extension
  public static final class DescriptorImpl extends ToolInstallerDescriptor<QualityCenterClientInstaller> {

    public String getDisplayName() {
      return Messages.QualityCenterClientInstaller_DescriptorImpl_DisplayName();
    }

    @Override
    public boolean isApplicable(Class<? extends ToolInstallation> toolType) {
      return toolType==QualityCenterClientInstallation.class;
    }

    public FormValidation doCheckQcServerURL(@QueryParameter String value) {
      // this can be used to check the existence of a file on the server, so needs to be protected
      Hudson.getInstance().checkPermission(Hudson.ADMINISTER);
      return QualityCenterUtils.checkQcServerURL(value, true);
    }

    public FormValidation doCheckLocalPathToQCClient(@QueryParameter String value) {
      // this can be used to check the existence of a file on the server, so needs to be protected
      Hudson.getInstance().checkPermission(Hudson.ADMINISTER);
      return QualityCenterUtils.checkLocalPathToInstaller(value, true);
    }

  }

}
