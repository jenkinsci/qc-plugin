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

package com.michelin.cio.hudson.plugins.qc.qtpaddins;

import com.michelin.cio.hudson.plugins.qc.client.QualityCenterClientInstallation;
import com.michelin.cio.hudson.plugins.qc.Messages;
import com.michelin.cio.hudson.plugins.qc.QualityCenterUtils;
import groovy.text.GStringTemplateEngine;
import hudson.AbortException;
import hudson.Extension;
import hudson.FilePath;
import hudson.Launcher;
import hudson.ProxyConfiguration;
import hudson.model.Hudson;
import hudson.model.Node;
import hudson.model.TaskListener;
import hudson.tools.ToolInstallation;
import hudson.tools.ToolInstaller;
import hudson.tools.ToolInstallerDescriptor;
import hudson.util.ArgumentListBuilder;
import hudson.util.FormValidation;
import java.io.File;
import java.io.IOException;
import java.io.PrintStream;

import java.net.MalformedURLException;
import java.net.URISyntaxException;
import java.net.URL;
import java.net.URLConnection;
import java.util.HashMap;
import java.util.Map;
import org.kohsuke.stapler.DataBoundConstructor;
import org.kohsuke.stapler.QueryParameter;


/**
 * This class represents the installer for QuickTest Professional Add-in.
 *
 * <p>We provide two ways to get the installer:<ol>
 * <li>The first one consists in downloading the right installer from <a
 * href="http://update.external.hp.com/qualitycenter/qc90/mictools/qtp/index.html">
 * HP Update Center</a>, so an Internet connection is required (don't forget to
 * set the proxy parameters in Hudson if applicable);</li>
 * <li>The second one consists in getting the installer from a directory on the
 * Hudson master node's filesystem.</li>
 * </ol></p>
 *
 * @author Thomas Maurel
 */
public class QualityCenterQTPAddinsInstaller extends ToolInstaller {

  private static final String TEMPLATE_NAME = "installTemplate.iss";
  private static final String GENERATED_ISS_NAME = "setup.iss";
  private static final String REPORT_READER_EXE = "QTReport.exe";
  private static final String BIN_FOLDER = "bin";

  private final String version;
  private final String localPathToQTPAddin;
  private final boolean acceptLicense;

  @DataBoundConstructor
  public QualityCenterQTPAddinsInstaller(String version, String localPathToQTPAddin, boolean acceptLicense) {
      super(null);
      this.version = version;
      this.localPathToQTPAddin = localPathToQTPAddin;
      this.acceptLicense = acceptLicense;
  }

  public boolean isAcceptLicense() {
    return acceptLicense;
  }

  public String getLocalPathToQTPAddin() {
    return localPathToQTPAddin;
  }

  public String getVersion() {
    return version;
  }

  @Override
  public FilePath performInstallation(ToolInstallation tool, Node node, TaskListener log) throws InterruptedException, AbortException, IOException  {
    FilePath expectedLocation = preferredLocation(tool, node);
    PrintStream out = log.getLogger();

    // Check if has been installed by hudson or manually
    if(expectedLocation.child(".installedByHudson").exists() || expectedLocation.child(BIN_FOLDER).child(REPORT_READER_EXE).exists()) {
      return expectedLocation;
    }

    // Create the directory tree
    expectedLocation.mkdirs();
    FilePath file = expectedLocation.child(GENERATED_ISS_NAME);

    // Did he accept the license agreement?
    if(!acceptLicense) {
      log.fatalError(Messages.QualityCenterQTPAddinsInstaller_AcceptLicense());
      throw new AbortException();
    }

    // Get the URL to the bundled InstallShield silent install script template
    URL template = Hudson.getInstance().pluginManager.uberClassLoader.getResource(TEMPLATE_NAME);
    File f;
    try {
      f = new File(template.toURI());
    } catch(URISyntaxException e) {
      f = new File(template.getPath());
    }

    // Note from the code reviewer:  The following could clearly have been done
    // with Velocity (and it would surely have been faster), but, well, let's
    // let the kids play a little bit!
    //
    // Groovy fun
    // To perform a silent installation, we need to have a setup.iss file in the same
    // folder as the installer which will describe each step of the installation.
    // Though we need to specify the installer key for each step, and the key varies for each
    // version which is why we need to parse a template.
    GStringTemplateEngine engine = new GStringTemplateEngine();
    // Get the defined version of the QTPAddin
    QTPVersion currentVersion = QTPVersion.valueOf("QTP" + this.version.replaceAll("\\.", ""));
    if(version == null) {
      log.fatalError(Messages.QualityCenterQTPAddinsInstaller_CouldntFindValidVersion());
      throw new AbortException();
    }

    // Build a map to parse the template
    Map<String, String> binding = new HashMap<String, String>();
    binding.put("key", currentVersion.key);
    binding.put("path", expectedLocation.absolutize().getRemote());

    String instalIss;
    out.println(Messages.QualityCenterQTPAddinsInstaller_GeneratingInstallerISS());
    try {
      // Parse the template and put the result in a string
      instalIss = engine.createTemplate(f).make(binding).toString();
    } catch (Exception e) {
      log.fatalError(Messages.QualityCenterQTPAddinsInstaller_CouldntGenerateInstallerISS());
      throw new AbortException();
    }

    // Put the parsed template in the project workspace
    file.write(instalIss, "ISO-8859-1");

    FilePath installer;

    // If the installer is stocked on master, copy it from there
    if(localPathToQTPAddin != null & localPathToQTPAddin.length() > 0) {
      FilePath installerOnMaster = new FilePath(Hudson.MasterComputer.localChannel, localPathToQTPAddin);
      if(!installerOnMaster.exists()) {
        log.fatalError(Messages.QualityCenterClientInstaller_CannotFindInstaller());
        throw new AbortException();
      }
      
      if(installerOnMaster.isDirectory()) {
        log.fatalError(Messages.QualityCenterClientInstaller_ShouldBeAFile());
        throw new AbortException();
      }
      installer = expectedLocation.child(installerOnMaster.getName());
      out.println(Messages.QualityCenterClientInstaller_CopyingFromMaster(localPathToQTPAddin));
      installer.copyFrom(installerOnMaster);
    }
    // Else, download it from HP Update Center
    else {
      URL installURL = new URL(currentVersion.url);
      URLConnection cnx = ProxyConfiguration.open(installURL);
      installer = expectedLocation.child(installURL.getFile());
      out.println(Messages.QualityCenterClientInstaller_Downloading(installURL));
      installer.copyFrom(cnx.getInputStream());
    }

    // Perform install
    install(node.createLauncher(log), log, expectedLocation.absolutize().getRemote(), installer);

    // Successfully installed
    // Delete installer
    installer.delete();
    // Delete silent install script
    file.delete();

    // Does the bin folder exist after install ?
    if(!expectedLocation.child(BIN_FOLDER).child(REPORT_READER_EXE).exists()) {
      log.fatalError(Messages.QualityCenterQTPAddinsInstaller_CouldntFindExeAfterInstall(REPORT_READER_EXE));
      throw new AbortException();
    }
    expectedLocation.child(".installedByHudson").touch(System.currentTimeMillis());

    return expectedLocation;
  }

  public void install(Launcher launcher, TaskListener log, String expectedLocation, FilePath install) throws IOException, InterruptedException {
    PrintStream out = log.getLogger();
    String installName = install.getName();
    FilePath absolutizedInstall = install.absolutize();

    out.println(Messages.QualityCenterClientInstaller_Installing(installName));
    FilePath logFilePath = absolutizedInstall.getParent().child("setup.log");

    // Perform the silent install
    ArgumentListBuilder args = new ArgumentListBuilder();
    args.add(absolutizedInstall.getRemote());
    args.add("/S");
    args.add("/v/qn");

    if(launcher.launch().cmds(args).stdout(out).pwd(expectedLocation).join() != 0) {
        log.fatalError(Messages.QualityCenterClientInstaller_AbortedInstall());
        out.println(logFilePath.readToString());
        throw new AbortException();
    }

    out.println(Messages.QualityCenterQTPAddinsInstaller_InstallationSuccessfull());
  }

  @Extension
  public static final class DescriptorImpl extends ToolInstallerDescriptor<QualityCenterQTPAddinsInstaller> {

    public String getDisplayName() {
      return Messages.QualityCenterQTPAddinsInstaller_DescriptorImpl_DisplayName();
    }

    /**
     * Gets the supported versions of the QTP Addin.
     */
    public QTPVersion[] getAddinsArray() throws MalformedURLException {
      return QTPVersion.values();
    }

    @Override
    public boolean isApplicable(Class<? extends ToolInstallation> toolType) {
      return toolType==QualityCenterClientInstallation.class;
    }

    public FormValidation doCheckLocalPathToQTPAddin(@QueryParameter String value) {
      // This can be used to check the existence of a file on the server, so needs to be protected
      Hudson.getInstance().checkPermission(Hudson.ADMINISTER);
      return QualityCenterUtils.checkLocalPathToInstaller(value, true);
    }

  }

  /**
   * Currently supported versions of the QTP Addin.
   */
  public enum QTPVersion {

    QTP90(
      "9.0",
      "http://update.external.hp.com/qualitycenter/qc90/mictools/qtp/TDPlugInsSetup.exe",
      "B7677BB4-7E32-4430-90AE-E37EE6FED55E"
    ),

    QTP91(
      "9.1",
      "http://update.external.hp.com/qualitycenter/qc90/mictools/qtp/qtp_9_1/TDPlugInsSetup.exe",
      "B7677BB4-7E32-4430-90AE-E37EE6FED55E"
    ),

    QTP92(
      "9.2",
      "http://update.external.hp.com/qualitycenter/qc90/mictools/qtp/qtp_sp/TDPlugInsSetup.exe",
      "051B35E4-5986-4AD5-9470-693558044699"
    );

    public final String version;

    /**
     * URL to download the addin from HP Update Center.
     */
    public final String url;

    /**
     * The key is the ID of the installer used to generate the InstallShield silent install script.
     */
    public final String key;

    QTPVersion(String version, String url, String key) {
      this.version = version;
      this.url = url;
      this.key = key;
    }

    public String getKey() {
      return key;
    }

    public String getUrl() {
      return url;
    }

    public String getVersion() {
      return version;
    }

  }
  
}
