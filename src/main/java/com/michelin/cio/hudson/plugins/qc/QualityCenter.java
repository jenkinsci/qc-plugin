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

package com.michelin.cio.hudson.plugins.qc;

import com.michelin.cio.hudson.plugins.qc.client.QualityCenterClientInstallation;
import com.michelin.cio.hudson.plugins.qc.qtpaddins.QualityCenterQTPAddinsInstallation;
import hudson.AbortException;
import hudson.CopyOnWrite;
import hudson.EnvVars;
import hudson.Extension;
import hudson.FilePath;
import hudson.Launcher;
import hudson.Util;
import hudson.model.AbstractBuild;
import hudson.model.AbstractProject;
import hudson.model.BuildListener;
import hudson.model.Computer;
import hudson.model.Hudson;
import hudson.tasks.BuildStepDescriptor;
import hudson.tasks.Builder;
import hudson.tools.ToolInstallation;
import hudson.util.ArgumentListBuilder;
import hudson.util.FormValidation;
import hudson.util.VariableResolver;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;
import java.io.PrintStream;
import java.net.URL;
import net.sf.json.JSONObject;
import org.apache.commons.io.IOUtils;
import org.apache.commons.lang.StringUtils;
import org.kohsuke.stapler.DataBoundConstructor;
import org.kohsuke.stapler.QueryParameter;
import org.kohsuke.stapler.StaplerRequest;

/**
 * {@link BuildStep} to run a TestSet from a Quality Center server.
 * 
 * @author Thomas Maurel
 */
public class QualityCenter extends Builder {

  private final static String VB_SCRIPT_NAME = "runTestSet.vbs";

  /**
   * Quality Center installation name.
   */
  private final String qcClientInstallationName;

  /**
   * QTP Addin for Quality Center installation name.
   */
  private final String qcQTPAddinInstallationName;

  /**
   * Quality Center server URL.
   */
  private final String qcServerURL;

  /**
   * Username to log into the Quality Center server.
   */
  private final String qcLogin;

  /**
   * The password associated to the username to log into the Quality Center server.
   */
  private final String qcPass;

  /**
   * The domain where the Quality Center project is located.
   */
  private final String qcDomain;

  /**
   * The Quality Center project to connect to.
   */
  private final String qcProject;

  /**
   * The folder where the TestSet to run is located.
   */
  private final String qcTSFolder;

  /**
   * The name of the TestSet to run.
   */
  private final String qcTSName;

  /**
   * The name of the report file.
   */
  private final String qcTSLogFile;

  /**
   * Timeout
   */
  private final String qcTimeOut;
  
  /**
   * The parsed name (env vars) of the report file.
   */
  private String parsedQcTSLogFile;


  @DataBoundConstructor
  public QualityCenter(
            String qcClientInstallationName,
            String qcQTPAddinInstallationName,
            String qcServerURL,
            String qcLogin,
            String qcPass,
            String qcDomain,
            String qcProject,
            String qcTSFolder,
            String qcTSName,
            String qcTSLogFile,
            String qcTimeOut) {
    this.qcClientInstallationName = qcClientInstallationName;
    this.qcQTPAddinInstallationName = qcQTPAddinInstallationName;
    this.qcServerURL = qcServerURL;
    this.qcLogin = qcLogin;
    this.qcPass = qcPass;
    this.qcDomain = qcDomain;
    this.qcProject = qcProject;
    this.qcTSFolder = qcTSFolder;
    this.qcTSName = qcTSName;
    this.qcTSLogFile = qcTSLogFile;
    this.qcTimeOut = qcTimeOut;
  }

  public String getQcDomain() {
    return qcDomain;
  }

  public String getQcClientInstallationName() {
    return qcClientInstallationName;
  }

  public String getQcQTPAddinInstallationName() {
    return qcQTPAddinInstallationName;
  }

  public String getQcLogin() {
    return qcLogin;
  }

  public String getQcPass() {
    return qcPass;
  }

  public String getQcProject() {
    return qcProject;
  }

  public String getQcServerURL() {
    return qcServerURL;
  }

  public String getQcTSFolder() {
    return qcTSFolder;
  }

  public String getQcTSLogFile() {
    return qcTSLogFile;
  }

  public String getParsedQcTSLogFile() {
    return parsedQcTSLogFile;
  }

  public String getQcTSName() {
    return qcTSName;
  }

  public String getQcTimeOut() {
    return qcTimeOut;
  }

  public static String getVbScriptName() {
    return VB_SCRIPT_NAME;
  }

  @Override
  public DescriptorImpl getDescriptor() {
      return (DescriptorImpl) super.getDescriptor();
  }

  public QualityCenterClientInstallation getQualityCenterClientInstallation() {
    for(QualityCenterClientInstallation installation: getDescriptor().getClientInstallations()) {
      if(this.qcClientInstallationName != null && installation.getName().equals(this.qcClientInstallationName)) {
        return installation;
      }
    }
    return null;
  }

  public QualityCenterQTPAddinsInstallation getQualityCenterQTPAddinInstallation() {
    for(QualityCenterQTPAddinsInstallation installation: getDescriptor().getQTPAddinsInstallations()) {
      if(this.qcQTPAddinInstallationName != null && installation.getName().equals(this.qcQTPAddinInstallationName)) {
        return installation;
      }
    }
    return null;
  }

  @Override
  public boolean perform(AbstractBuild<?, ?> build, Launcher launcher, BuildListener listener) throws InterruptedException, IOException {
    EnvVars env = build.getEnvironment(listener);

    // Has a QC installation been set? If yes, is it really a QC installation?
    QualityCenterClientInstallation qcInstallation = getQualityCenterClientInstallation();
    if(qcInstallation == null) {
      // No installation has been set
      listener.fatalError(Messages.QualityCenter_NoInstallationSet());
      return false;
    }
    else {
      // Get an installation instance for this specific node.
      // Will run the QC Client auto-installer if not installed on this node
      qcInstallation = qcInstallation.forNode(Computer.currentComputer().getNode(), listener);
      // Will be null if not on a Windows Node
      if(qcInstallation == null) {
        listener.fatalError(Messages.QualityCenter_NotAvailableOnThisOS());
        return false;
      }

      qcInstallation = qcInstallation.forEnvironment(env);
      String qcDll = qcInstallation.getQCDll(launcher);
      // If we cant find the OTAClient DLL, then we can't run the testSet
      if(qcDll == null) {
        listener.fatalError(Messages.QualityCenter_DllNotFound());
        return false;
      }

      // Get an installation instance of the QTP Addin
      // QTP Addin must be installed but we dont need to call it directly.
      // It means defining QTPAddinsInstallations is only useful to auto-install it on nodes
      QualityCenterQTPAddinsInstallation qcQTPInstallation = getQualityCenterQTPAddinInstallation();

      if(qcQTPInstallation != null) {
        // Get the QTPAddin for this node, and install it if necessary
        qcQTPInstallation = qcQTPInstallation.forNode(Computer.currentComputer().getNode(), listener);
        qcQTPInstallation = qcQTPInstallation.forEnvironment(env);
      }

      FilePath projectWS = build.getWorkspace();
      // Get the URL to the VBScript used to run the test, which is bundled in the plugin
      URL vbsUrl = Hudson.getInstance().pluginManager.uberClassLoader.getResource(VB_SCRIPT_NAME);
      if(vbsUrl == null) {
        listener.fatalError(Messages.QualityCenter_VBSNotFound());
        return false;
      }

      // Copy the script to the project workspace
      FilePath vbScript = projectWS.child(VB_SCRIPT_NAME);
      vbScript.copyFrom(vbsUrl);

      try {
        // Run the VBScript
        String logFile = runVBScript(build, launcher, listener, vbScript);
        // Has the report been successfuly generated?
        if(!projectWS.child(logFile).exists()) {
          listener.fatalError(Messages.QualityCenter_ReportNotGenerated());
          return false;
        }
      }
      catch(IOException ioe) {
        Util.displayIOException(ioe, listener);
        return false;
      }
      finally {
        // Remove the VBScript from workspace
        vbScript.delete();
      }
    }

    return true;
  }

  /**
   * Adds some {@link EnvVars} to the project's {@link EnvVars} table.
   * 
   * <p>This {@link EnvVars} can be used in the report file name.</p>
   */
  private void pushEnvVars(EnvVars env) {
    env.put("QC_DOMAIN", this.qcDomain);
    env.put("QC_PROJECT", this.qcProject);
    env.put("TS_FOLDER", this.qcTSFolder);
    env.put("TS_NAME", this.qcTSName);
  }

  /**
   * Removes the {@link EnvVars} from the project's {@link EnvVars} table.
   */
  private void removeEnvVars(EnvVars env) {
    env.remove("QC_DOMAIN");
    env.remove("QC_PROJECT");
    env.remove("TS_FOLDER");
    env.remove("TS_NAME");
  }

  /**
   * Runs the TestSet through VBScript.
   */
  private String runVBScript(AbstractBuild<?, ?> build, Launcher launcher, BuildListener listener, FilePath file) throws IOException, InterruptedException {
    ArgumentListBuilder args = new ArgumentListBuilder();
    EnvVars env = build.getEnvironment(listener);
    VariableResolver<String> varResolver = build.getBuildVariableResolver();
    PrintStream out = listener.getLogger();

    // Add the qc specific env vars
    pushEnvVars(env);

    // Parse the report file name using env vars
    this.parsedQcTSLogFile = Util.replaceMacro(env.expand(this.qcTSLogFile), varResolver);

    // Use cscript to run the vbscript and get the console output
    args.add("cscript");
    args.add("/nologo");
    args.add(file);
    args.add(Util.replaceMacro(env.expand(this.qcServerURL), varResolver));
    args.add(Util.replaceMacro(env.expand(this.qcLogin), varResolver));

    // If no password, then replace by ""
    if(!StringUtils.isBlank(this.qcPass)) {
      args.addMasked(Util.replaceMacro(env.expand(this.qcPass), varResolver));
    }
    else {
      args.addMasked("\"\"");
    }
    args.add(Util.replaceMacro(env.expand(this.qcDomain), varResolver));
    args.add(Util.replaceMacro(env.expand(this.qcProject), varResolver));
    args.add(Util.replaceMacro(env.expand(this.qcTSFolder), varResolver));
    args.add(Util.replaceMacro(env.expand(this.qcTSName), varResolver));
    args.add(this.parsedQcTSLogFile);
    args.add(this.qcTimeOut);

    // Remove qc specific environment variables
    removeEnvVars(env);

    // Run the script on node
    // Execution result should be 0
    if(launcher.launch().cmds(args).stdout(out).pwd(file.getParent()).join() != 0) {
      listener.fatalError(Messages.QualityCenter_TSSchedulerFailed());

      // log file is in UTF-16
      InputStream is= new FileInputStream(this.parsedQcTSLogFile);
      InputStreamReader in = new InputStreamReader(is, "UTF-16");
      try {
          // Copy the console output to our logger
          IOUtils.copy(in,new OutputStreamWriter(out));
      } finally {
          in.close();
          is.close();
      }
      throw new AbortException();
    }

    return this.parsedQcTSLogFile;
  }

  @Extension
  public static class DescriptorImpl extends BuildStepDescriptor<Builder> {

    /**
     * Installations of the QCClient
     */
    @CopyOnWrite
    private volatile QualityCenterClientInstallation[] clientInstallations = new QualityCenterClientInstallation[0];

    /**
     * Installations of the QTP Addin
     */
    @CopyOnWrite
    private volatile QualityCenterQTPAddinsInstallation[] qtpAddinsInstallations = new QualityCenterQTPAddinsInstallation[0];

    public DescriptorImpl() {
      load();
    }

    protected DescriptorImpl(Class<? extends QualityCenter> clazz) {
      super(clazz);
    }

    @Override
    public String getDisplayName() {
      return Messages.QualityCenter_DisplayName();
    }

    public QualityCenterClientInstallation.DescriptorImpl getToolDescriptor() {
      return ToolInstallation.all().get(QualityCenterClientInstallation.DescriptorImpl.class);
    }

    @Override
    public boolean isApplicable(Class<? extends AbstractProject> jobType) {
      if(getClientInstallations() != null && getClientInstallations().length > 0 &&
          getQTPAddinsInstallations() != null && getQTPAddinsInstallations().length > 0) {
        return true;
      }
      return false;
    }

    @Override
    public Builder newInstance(StaplerRequest req, JSONObject formData) throws FormException {
      return req.bindJSON(QualityCenter.class, formData);
    }

    public QualityCenterClientInstallation[] getClientInstallations() {
      return clientInstallations;
    }

    public void setClientInstallations(QualityCenterClientInstallation... installations) {
      this.clientInstallations = installations;
      save();
    }

    public QualityCenterQTPAddinsInstallation[] getQTPAddinsInstallations() {
      return qtpAddinsInstallations;
    }

    public void setQTPAddinsInstallations(QualityCenterQTPAddinsInstallation... installations) {
      this.qtpAddinsInstallations = installations;
      save();
    }

    public FormValidation doCheckQcServerURL(@QueryParameter String value) {
      return QualityCenterUtils.checkQcServerURL(value);
    }

    public FormValidation doCheckQcLogin(@QueryParameter String value) {
      if(StringUtils.isBlank(value)) {
        return FormValidation.error(Messages.QualityCenter_UsernameShouldBeDefined());
      }
      return FormValidation.ok();
    }

    public FormValidation doCheckQcDomain(@QueryParameter String value) {
      if(StringUtils.isBlank(value)) {
        return FormValidation.error(Messages.QualityCenter_DomainShouldBeDefined());
      }
      return FormValidation.ok();
    }

    public FormValidation doCheckQcProject(@QueryParameter String value) {
      if(StringUtils.isBlank(value)) {
        return FormValidation.error(Messages.QualityCenter_ProjectShouldBeDefined());
      }
      return FormValidation.ok();
    }

    public FormValidation doCheckQcTSFolder(@QueryParameter String value) {
      if(StringUtils.isBlank(value)) {
        return FormValidation.error(Messages.QualityCenter_TSFolderShouldBeDefined());
      }
      return FormValidation.ok();
    }

    public FormValidation doCheckQcTSName(@QueryParameter String value) {
      if(StringUtils.isBlank(value)) {
        return FormValidation.error(Messages.QualityCenter_TSNameShouldBeDefined());
      }
      return FormValidation.ok();
    }

  }

}
