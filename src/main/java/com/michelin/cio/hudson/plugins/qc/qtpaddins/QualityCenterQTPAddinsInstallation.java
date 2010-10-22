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

import com.michelin.cio.hudson.plugins.qc.Messages;
import com.michelin.cio.hudson.plugins.qc.QualityCenter;
import hudson.EnvVars;
import hudson.Extension;
import hudson.Functions;
import hudson.model.Computer;
import hudson.model.EnvironmentSpecific;
import hudson.model.Hudson;
import hudson.model.Node;
import hudson.model.TaskListener;
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
 * This class represents a QuickTest Professional Add-in installation.
 *
 * <p>As specified on <a href="http://update.external.hp.com/qualitycenter/qc90/mictools/qtp/index.html">
 * HP's Web site</a>, "QuickTest Professional Add-in integrates Quality Center
 * with QuickTest Professional". It means that to run Quality Center TestSets,
 * it is required to install two components:<ol>
 * <li>The Quality Center client, represented by {@link QualityCenterClientInstallation};</li>
 * <li>The QuickTest Professional Add-in installation, represented by this class.</li>
 * </ol></p>
 *
 * @author Thomas Maurel
 */
public class QualityCenterQTPAddinsInstallation extends ToolInstallation implements NodeSpecific<QualityCenterQTPAddinsInstallation>, EnvironmentSpecific<QualityCenterQTPAddinsInstallation> {

  public QualityCenterQTPAddinsInstallation(String name, String qcHome) {
    super(name, qcHome, Collections.<ToolProperty<?>>emptyList());
  }

  @DataBoundConstructor
  public QualityCenterQTPAddinsInstallation(String name, String home, List<? extends ToolProperty<?>> properties) {
    super(name, home, properties);
  }

  /**
   * Gets the installation for the specified node.
   * @return {@code null} if not a Windows node
   */
  public QualityCenterQTPAddinsInstallation forNode(Node node, TaskListener log) throws IOException, InterruptedException {
    // Is the job running on a slave node?
    Computer computer = node.toComputer();
    if(computer instanceof SlaveComputer && !((SlaveComputer) computer).isUnix()) {
        computer = null;
        return new QualityCenterQTPAddinsInstallation(getName(), translateFor(node, log));
    }

    // Is the job running on the master node?
    if(Functions.isWindows()) {
      return new QualityCenterQTPAddinsInstallation(getName(), translateFor(node, log));
    }

    return null;
  }

  public QualityCenterQTPAddinsInstallation forEnvironment(EnvVars environment) {
    return new QualityCenterQTPAddinsInstallation(getName(), environment.expand(getHome()));
  }

  @Extension
  public static class DescriptorImpl extends ToolDescriptor<QualityCenterQTPAddinsInstallation> {

    public String getDisplayName() {
      return Messages.QualityCenterQTPAddinsInstallation_DisplayName();
    }

    @Override
    public List<QualityCenterQTPAddinsInstaller> getDefaultInstallers() {
      return Collections.singletonList(new QualityCenterQTPAddinsInstaller(null, null, false));
    }

    /**
     * Persisted by the {@link QualityCenter} descriptor.
     */
    @Override
    public QualityCenterQTPAddinsInstallation[] getInstallations() {
      return Hudson.getInstance().getDescriptorByType(QualityCenter.DescriptorImpl.class).getQTPAddinsInstallations();
    }

    @Override
    public void setInstallations(QualityCenterQTPAddinsInstallation... installations) {
      Hudson.getInstance().getDescriptorByType(QualityCenter.DescriptorImpl.class).setQTPAddinsInstallations(installations);
    }

    /**
     * Checks if QTP home is a valid QC home path.
     */
    public FormValidation doCheckHome(@QueryParameter File value) {
      Hudson.getInstance().checkPermission(Hudson.ADMINISTER);

      if(value.getPath().equals(""))
        return FormValidation.ok();

      if(!value.isDirectory())
        return FormValidation.error(Messages.QualityCenterInstallation_NotADirectory());

      return FormValidation.ok();
    }

  }
  
}
