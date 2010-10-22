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

import hudson.util.FormValidation;
import java.io.File;
import java.io.IOException;
import java.net.HttpURLConnection;
import java.net.MalformedURLException;
import java.net.URL;
import org.apache.commons.lang.StringUtils;

/**
 * @author Thomas Maurel
 */
public class QualityCenterUtils {

  public static FormValidation checkQcServerURL(String value) {
    return checkQcServerURL(value, false);
  }

  /**
   * Checks the Quality Center server URL.
   */
  public static FormValidation checkQcServerURL(String value, Boolean acceptEmpty) {
      String url;
      // Path to the page to check if the server is alive
      String page = "servlet/tdservlet/TDAPI_GeneralWebTreatment";

      // Do will allow empty value?
      if(StringUtils.isBlank(value)) {
        if(!acceptEmpty) {
          return FormValidation.error(Messages.QualityCenter_ServerURLMustBeDefined());
        }
        else {
          return FormValidation.ok();
        }
      }

      // Does the URL ends with a "/" ? if not, add it
      if(value.lastIndexOf("/") == value.length() - 1) {
        url = value + page;
      }
      else {
        url = value + "/" + page;
      }

      // Open the connection and perform a HEAD request
      HttpURLConnection connection;
      try {
        connection = (HttpURLConnection) new URL(url).openConnection();
        connection.setRequestMethod("HEAD");

        // Check the response code
        if(connection.getResponseCode() != HttpURLConnection.HTTP_OK) {
          return FormValidation.error(connection.getResponseMessage());
        }
      } catch (MalformedURLException ex) {
        // This is not a valid URL
        return FormValidation.error(Messages.QualityCenter_MalformedServerURL());
      } catch (IOException ex) {
        // Cant open connection to the server
        return FormValidation.error(Messages.QualityCenter_ErrorOpeningServerConnection());
      }

      return FormValidation.ok();
    }

    /**
     * Checks the path to the QCClient on master server.
     */
    public static FormValidation checkLocalPathToInstaller(String value, Boolean acceptEmpty) {
      // Do will allow empty value?
       if(StringUtils.isBlank(value)) {
        if(!acceptEmpty) {
          return FormValidation.error(Messages.QualityCenter_PathMustBeDefined());
        }
        else {
          return FormValidation.ok();
        }
      }
      
      File installer = new File(value);

      if(!installer.exists()) {
        return FormValidation.error(Messages.QualityCenterClientInstaller_CannotFindInstaller());
      }

      // Must be a file
      if(!installer.isFile()) {
        return FormValidation.error(Messages.QualityCenterClientInstaller_ShouldBeAFile());
      }

      return FormValidation.ok();
    }

}
