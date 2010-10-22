/*
 * The MIT License
 *
 * Copyright (c) 2010, Manufacture Française des Pneumatiques Michelin, Thomas Maurel
 * Copyright (c) 2004-2009, Sun Microsystems, Inc., Kohsuke Kawaguchi, Martin Eigenbrodt, Tom Huybrechts, Yahoo!, Inc.
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

import hudson.Extension;
import hudson.FilePath.FileCallable;
import hudson.Launcher;
import hudson.AbortException;
import hudson.matrix.MatrixAggregatable;
import hudson.matrix.MatrixAggregator;
import hudson.matrix.MatrixBuild;
import hudson.model.AbstractBuild;
import hudson.model.AbstractProject;
import hudson.model.Action;
import hudson.model.BuildListener;
import hudson.model.Project;
import hudson.model.Result;
import hudson.remoting.VirtualChannel;
import hudson.tasks.BuildStepDescriptor;
import hudson.tasks.BuildStepMonitor;
import hudson.tasks.Builder;
import hudson.tasks.Publisher;
import hudson.tasks.Recorder;
import hudson.tasks.junit.TestResult;
import hudson.tasks.junit.TestResultAction;
import hudson.tasks.test.TestResultAggregator;
import hudson.tasks.test.TestResultProjectAction;
import org.apache.tools.ant.DirectoryScanner;
import org.kohsuke.stapler.DataBoundConstructor;

import java.io.File;
import java.io.IOException;
import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

/**
 * This class is adapted from {@link JunitResultArchiver}; Only the {@code perform()}
 * method slightly differs.
 * 
 * @author Thomas Maurel
 */
public class QualityCenterResultArchiver extends Recorder implements Serializable, MatrixAggregatable {

  @DataBoundConstructor
  public QualityCenterResultArchiver() {
  }

  @Override
  public boolean perform(AbstractBuild build, Launcher launcher, BuildListener listener) throws InterruptedException, IOException {
      TestResultAction action;
      List<Builder> builders = ((Project)build.getProject()).getBuilders();
      final List<String> names = new ArrayList<String>();
      // Get the TestSet report files names of the current build
      for(Builder builder : builders) {
        if(builder instanceof QualityCenter) {
          names.add(((QualityCenter)builder).getParsedQcTSLogFile());
        }
      }

      // Has any QualityCenter builder been set up?
      if(names.isEmpty()) {
        listener.getLogger().println(Messages.QualityCenterResultArchiver_NoBuilderSet());
        return true;
      }

      try {
          final long buildTime = build.getTimestamp().getTimeInMillis();
          final long nowMaster = System.currentTimeMillis();

          TestResult result = build.getWorkspace().act(new FileCallable<TestResult>() {
              public TestResult invoke(File ws, VirtualChannel channel) throws IOException {
                  final long nowSlave = System.currentTimeMillis();
                  List<String> files = new ArrayList<String>();
                  DirectoryScanner ds = new DirectoryScanner();
                  ds.setBasedir(ws);

                  // Transform the report file names list to a File Array,
                  // and add it to the DirectoryScanner includes set
                  for(String name : names) {
                    
                    File file = new File(ws, name);
                    if(file.exists()) {
                      files.add(file.getName());
                    }
                  }

                  Object[] objectArray = new String[files.size()];
                  files.toArray(objectArray);
                  ds.setIncludes((String[])objectArray);
                  ds.scan();
                  if(ds.getIncludedFilesCount()==0) {
                      // no test result. Most likely a configuration error or fatal problem
                      throw new AbortException("Report not found");
                  }

                  return new TestResult(buildTime+(nowSlave-nowMaster), ds);
              }
          });

          action = new TestResultAction(build, result, listener);
          if(result.getPassCount()==0 && result.getFailCount()==0) {
              throw new AbortException("Result is empty");
          }
      } catch (AbortException e) {
          if(build.getResult()==Result.FAILURE) {
              // most likely a build failed before it gets to the test phase.
              // don't report confusing error message.
              return true;
          }

          listener.getLogger().println(e.getMessage());
          build.setResult(Result.FAILURE);
          return true;
      } catch (IOException e) {
          e.printStackTrace(listener.error("Failed to archive QC reports"));
          build.setResult(Result.FAILURE);
          return true;
      }

      build.getActions().add(action);

      if(action.getResult().getFailCount()>0) {
          build.setResult(Result.UNSTABLE);
      }

      return true;
  }

  @Override
  public Action getProjectAction(AbstractProject<?, ?> project) {
      return new TestResultProjectAction(project);
  }

  public MatrixAggregator createAggregator(MatrixBuild build, Launcher launcher, BuildListener listener) {
      return new TestResultAggregator(build,launcher,listener);
  }

  public BuildStepMonitor getRequiredMonitorService() {
      return BuildStepMonitor.BUILD;
  }

  @Extension
  public static class DescriptorImpl extends BuildStepDescriptor<Publisher> {

    public String getDisplayName() {
        return Messages.QualityCenterResultArchiver_DisplayName();
    }

    public boolean isApplicable(Class<? extends AbstractProject> jobType) {
        return true;
    }
    
  }

}
