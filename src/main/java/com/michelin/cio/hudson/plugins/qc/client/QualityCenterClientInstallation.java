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
import com.michelin.cio.hudson.plugins.qc.QualityCenter;
import hudson.EnvVars;
import hudson.Extension;
import hudson.Functions;
import hudson.Launcher;
import hudson.Util;
import hudson.model.Computer;
import hudson.model.EnvironmentSpecific;
import hudson.model.Hudson;
import hudson.model.Node;
import hudson.model.TaskListener;
import hudson.remoting.Callable;
import hudson.slaves.NodeSpecific;
import hudson.slaves.SlaveComputer;
import hudson.tools.ToolDescriptor;
import hudson.tools.ToolInstallation;
import hudson.tools.ToolProperty;
import hudson.util.FormValidation;
import java.io.File;
import java.io.IOException;
import java.util.Collections;
import java.util.List;
import org.kohsuke.stapler.DataBoundConstructor;
import org.kohsuke.stapler.QueryParameter;

/**
 * @author Thomas Maurel
 */
public class QualityCenterClientInstallation extends ToolInstallation implements NodeSpecific<QualityCenterClientInstallation>, EnvironmentSpecific<QualityCenterClientInstallation> {

  public QualityCenterClientInstallation(String name, String qcHome) {
    super(name, qcHome, Collections.<ToolProperty<?>>emptyList());
  }

  @DataBoundConstructor
  public QualityCenterClientInstallation(String name, String home, List<? extends ToolProperty<?>> properties) {
    super(name, home, properties);
  }

  /**
   * Gets the installation for the specified node.
   * @return {@code null} if not a Windows node
   */
  public QualityCenterClientInstallation forNode(Node node, TaskListener log) throws IOException, InterruptedException {
    // Is the job running on a slave node?
    Computer computer = node.toComputer();
    if(computer instanceof SlaveComputer && !((SlaveComputer) computer).isUnix()) {
        computer = null;
        return new QualityCenterClientInstallation(getName(), translateFor(node, log));
    }
    
    // Is the job running on the master node?
    if(Functions.isWindows()) {
      return new QualityCenterClientInstallation(getName(), translateFor(node, log));
    }

    return null;
  }

  public QualityCenterClientInstallation forEnvironment(EnvVars environment) {
    return new QualityCenterClientInstallation(getName(), environment.expand(getHome()));
  }

  /**
   * Returns the path to {@code OTAClient.dll}.
   */
  public String getQCDll(Launcher launcher) throws IOException, InterruptedException {
    return launcher.getChannel().call(new Callable<String,IOException>() {
      public String call() throws IOException {
        File qcDll = getDllFile();
        if(qcDll.exists()) {
          return qcDll.getPath();
        }
        return null;
      }
    });
  }

  private File getDllFile() {
    String home = Util.replaceMacro(getHome(), EnvVars.masterEnvVars);
    return new File(home, QualityCenterClientInstaller.DLL_NAME);
  }

  @Extension
  public static class DescriptorImpl extends ToolDescriptor<QualityCenterClientInstallation> {

    public String getDisplayName() {
      return Messages.QualityCenterInstallation_DisplayName();
    }

    @Override
    public List<QualityCenterClientInstaller> getDefaultInstallers() {
      return Collections.singletonList(new QualityCenterClientInstaller(null, null));
    }

    /**
     * Persisted by the {@link QualityCenter} descriptor.
     */
    @Override
    public QualityCenterClientInstallation[] getInstallations() {
      return Hudson.getInstance().getDescriptorByType(QualityCenter.DescriptorImpl.class).getClientInstallations();
    }

    @Override
    public void setInstallations(QualityCenterClientInstallation... installations) {
      Hudson.getInstance().getDescriptorByType(QualityCenter.DescriptorImpl.class).setClientInstallations(installations);
    }

    /**
     * Checks if QC home is a valid QC home path.
     */
    public FormValidation doCheckHome(@QueryParameter File value) {
      Hudson.getInstance().checkPermission(Hudson.ADMINISTER);

      if(value.getPath().equals(""))
        return FormValidation.ok();

      if(!value.isDirectory())
        return FormValidation.error(Messages.QualityCenterInstallation_NotADirectory());

      File dll = new File(value, QualityCenterClientInstaller.DLL_NAME);
      File dll2 = new File(value, QualityCenterClientInstaller.DLL_PATH_ON_NODE + QualityCenterClientInstaller.DLL_NAME);

      if(!dll.exists() && !dll2.exists())
        return FormValidation.error(Messages.QualityCenterInstallation_NotQualityCenterDir());

      return FormValidation.ok();
    }
    
  }
  
}
