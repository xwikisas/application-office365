/*
 * See the NOTICE file distributed with this work for additional
 * information regarding copyright ownership.
 *
 * This is free software; you can redistribute it and/or modify it
 * under the terms of the GNU Lesser General Public License as
 * published by the Free Software Foundation; either version 2.1 of
 * the License, or (at your option) any later version.
 *
 * This software is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this software; if not, write to the Free
 * Software Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA
 * 02110-1301 USA, or see the FSF site: http://www.fsf.org.
 */
package com.xwiki.microsoft365.internal;

import java.io.IOException;
import java.util.List;
import java.util.Map;

import javax.inject.Inject;
import javax.inject.Named;
import javax.inject.Provider;
import javax.inject.Singleton;

import org.slf4j.Logger;
import org.xwiki.component.annotation.Component;
import org.xwiki.component.phase.Initializable;
import org.xwiki.component.phase.InitializationException;
import org.xwiki.model.reference.DocumentReference;
import org.xwiki.rendering.macro.wikibridge.WikiMacroParameters;

import com.xpn.xwiki.XWikiContext;
import com.xpn.xwiki.doc.XWikiDocument;
import com.xpn.xwiki.objects.BaseObject;
import com.xwiki.identityoauth.IdentityOAuthException;
import com.xwiki.identityoauth.IdentityOAuthProvider;
import com.xwiki.microsoft365.Microsoft365Connections;

/**
 * The environment used to run the microsoft365 macros.
 *
 * @version $Id$
 * @since 1.0
 */
@Component(roles = MacroRunner.class)
@Singleton
public final class MacroRunner implements Initializable, Microsoft365Constants
{
    @Inject
    private Logger logger;

    @Inject
    private Provider<XWikiContext> contextProvider;

    @Inject
    @Named("microsoft365")
    private Provider<IdentityOAuthProvider> ms365Provider;

    private Microsoft365ConnectionsImpl msConnections;

    private DocumentReference embeddedDocClass;

    @Override public void initialize() throws InitializationException
    {
        embeddedDocClass = new DocumentReference(WIKI_NAME,
                SPACE_NAME, EMBED_DOC_CLASSNAME);
    }

    private void initMsConnections()
    {
        if (msConnections == null) {
            msConnections = (Microsoft365ConnectionsImpl) (ms365Provider.get());
        }
    }

    private int readMacroNumParam(Object macroParam, String reqParam)
    {
        String nbS = (String) macroParam;
        String emptyVelocity = "${nb}";
        int macroNum = -1;
        if (nbS != null && nbS.trim().length() > 0 && !emptyVelocity.equals(nbS)) {
            macroNum = Integer.parseInt(nbS);
        } else if (reqParam != null && reqParam.trim().length() > 0 && !emptyVelocity.equals(reqParam)) {
            macroNum = Integer.parseInt((reqParam));
        }

        return macroNum;
    }

    private int extractMacroPos()
    {
        String key = "ms365macroNum";
        XWikiContext ctx = contextProvider.get();
        if (ctx.containsKey(key)) {
            int o = (Integer) ctx.get(key);
            o = o + 1;
            ctx.put(key, o);
            return o;
        } else {
            ctx.put(key, 0);
            return 0;
        }
    }

    MacroRun runMacro(Map<String, Object> macro)
    {
        initMsConnections();
        WikiMacroParameters macroParams = (WikiMacroParameters) macro.get("parameters");
        MacroRunner.MacroRun macroRun = new MacroRun(macroParams);

        int macroPos = extractMacroPos();
        //the macro-number given in parameter
        String nb = "nb";
        int macroNumParam =
                readMacroNumParam(macroParams.get(nb), contextProvider.get().getRequest().getParameter(nb));
        macroRun.setNumber(macroPos);

        // is there a request parameter that means writeObject? Then do so
        String writeObject = (String) contextProvider.get().getRequest().getParameter("writeObject");

        try {
            BaseObject obj = getEmbedObject(macroPos);
            if (writeObject != null && "do".equals(writeObject) && macroPos == macroNumParam) {
                performSave(macroPos, obj);
                // which also redirects
                macroRun.mode = "saveDocumentChoice";
                return macroRun;
            }

            String action = contextProvider.get().getAction();
            String filename = obj == null ? null : obj.getStringValue(FILENAME).trim();

            // from here we deliver content to be displayed
            if (filename != null && filename.length() > 0 && VIEW.equals(action)) {
                prepareDisplayEmbed(macroRun, obj);
            }
        } catch (Exception e) {
            e.printStackTrace();
            macroRun.error = e.toString();
            macroRun.errorMessage = e.getMessage();
        }
        return macroRun;
    }

    private void prepareDisplayEmbed(MacroRun macroRun, BaseObject obj) throws IOException
    {
        // display the embedded document
        String site = obj.getStringValue(SITE);
        String fileId = obj.getStringValue(ID);
        if (obj.getStringValue(FILENAME).toLowerCase().endsWith(DOT_PDF)) {
            // make the URL the download-url
            macroRun.mode = MODE_DISPLAY_PDF;
            String driveURL = msConnections.getDriveURL(obj.getStringValue(SITE));
            String docURL = driveURL + "items/" + obj.getStringValue(ID);
            Map m = msConnections.getGraphApiInfo(obj.getStringValue(SITE), docURL);
            macroRun.url = (String) m.get("@microsoft.graph.downloadUrl");
        } else {
            macroRun.mode = MODE_DISPLAY_EMBED_IFRAME;
            macroRun.url = obj.getStringValue(EMBEDLINK);

            // THINKME: rather use this to get embed URLs?
            // msConnections.requestEmbedUrl(obj.getStringValue(ID), obj.getStringValue(SITE));
            // (but it needs a post and requires an MS identity according to
            // https://docs.microsoft.com/en-us/graph/api/driveitem-preview?view=graph-rest-1.0 )
        }
    }

    private BaseObject getEmbedObject(int macroNum)
    {
        List<BaseObject> objects = contextProvider.get().getDoc().getXObjects(embeddedDocClass);
        for (BaseObject obj : objects) {
            if (obj.getIntValue(MACRO_NUM, -1) == macroNum) {
                return obj;
            }
        }
        return null;
    }

    private void performSave(int macroNum, BaseObject objP)
    {
        try {
            XWikiContext ctx = contextProvider.get();
            XWikiDocument doc = ctx.getDoc();
            BaseObject obj = objP;

            if (obj == null) {
                int objNum = contextProvider.get().getDoc().createXObject(embeddedDocClass, ctx);
                obj = doc.getXObject(embeddedDocClass, objNum);
                obj.set(MACRO_NUM, Long.valueOf(macroNum), ctx);
            }

            Map<String, String[]> reqParam = contextProvider.get().getRequest().getParameterMap();

            if (reqParam.get(ERASE) != null && "true".equals(reqParam.get("erase")[0])) {
                obj.set(ID, "", ctx);
                obj.set(EMBEDLINK, "", ctx);
                obj.set(EDIT_LINK, "", ctx);
                obj.set(SITE, "", ctx);
                obj.set(VERSION, "", ctx);
                obj.set(FILENAME, "", ctx);
                ctx.getWiki().saveDocument(doc, "Removing Microsoft365 Document Embed", ctx);
            } else {
                obj.set(ID, reqParam.get(ID)[0], ctx);
                setParamIfNotEmpty(obj, EMBEDLINK, reqParam.get(EMBEDLINK)[0], ctx);
                setParamIfNotEmpty(obj, EDIT_LINK, reqParam.get(EDIT_LINK)[0], ctx);

                if (reqParam.get(SITE) != null && reqParam.get(SITE).length > 0 && reqParam.get(SITE)[0] != null) {
                    obj.set(SITE, reqParam.get(SITE)[0], ctx);
                }
                if (reqParam.get(VERSION) != null) {
                    obj.set(VERSION, reqParam.get(VERSION)[0], ctx);
                }
                obj.set(FILENAME, reqParam.get(FILENAME)[0], ctx);
                obj.set("user", ctx.getUserReference().getName(), ctx);
                ctx.getWiki().saveDocument(doc, "Inserting Microsoft365 Document Embed", ctx);
            }

            ctx.getResponse().sendRedirect(getViewPath());
        } catch (Exception e) {
            e.printStackTrace();
            throw new IdentityOAuthException("Error at saving macro object.", e);
        }
    }

    private void setParamIfNotEmpty(BaseObject obj, String name, String value, XWikiContext ctx)
    {
        if (value != null && value.length() > 0) {
            obj.set(name, value, ctx);
        }
    }

    private String getViewPath()
    {
        String path = contextProvider.get().getRequest().getPathInfo();
        if (path != null && path.length() > 0) {
            path = path.substring(path.lastIndexOf('/'));
        } else {
            path = ".";
        }
        return path;
    }

    class MacroRun implements Microsoft365Connections.MacroRun
    {
        private boolean authenticationNeeded;

        private boolean redirecting;

        private String error;

        private String errorMessage;

        private String width;

        private String height;

        private String url;

        private String mode;

        private int number;

        MacroRun(WikiMacroParameters macroParams)
        {
            for (String param : macroParams.getParameterNames()) {
                logger.warn(" -- param " + param + " ; " + macroParams.get(param)
                        + " (class " + macroParams.get(param).getClass() + ")");
            }
            width = (String) macroParams.get("width");
            height = (String) macroParams.get("height");
            url = (String) macroParams.get("url");
            authenticationNeeded = msConnections.isMissingAuth();

            if (this.url == null) {
                if (authenticationNeeded) {
                    mode = MODE_AUTHENTICATION_NEEDED;
                } else {
                    mode = MODE_DISPLAY_SEARCH;
                }
            } else {
                if (Boolean.TRUE.equals(macroParams.get("requireAuth"))) {
                    mode = MODE_AUTHENTICATION_NEEDED;
                } else if (url.toLowerCase().endsWith(".pdf")) {
                    mode = MODE_DISPLAY_PDF;
                } else {
                    mode = MODE_DISPLAY_EMBED_IFRAME;
                }
            }

            if (authenticationNeeded) {
                url = msConnections.getOAuthStartUrl();
                if (Boolean.parseBoolean((String) macroParams.get("authentication"))) {
                    try {
                        redirecting = true;
                        contextProvider.get().getResponse().sendRedirect(this.url);
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                } else {
                    redirecting = false;
                }
            }
        }

        @Override
        public String getMode()
        {
            return mode;
        }

        @Override
        public boolean isAuthenticationNeeded()
        {
            return authenticationNeeded;
        }

        @Override
        public boolean isRedirecting()
        {
            return redirecting;
        }

        @Override
        public String getUrl()
        {
            return url;
        }

        @Override
        public String getWidth()
        {
            return width;
        }

        @Override
        public String getHeight()
        {
            return height;
        }

        @Override
        public String getError()
        {
            return error;
        }

        @Override
        public String getErrorMessage()
        {
            return errorMessage;
        }

        @Override public int getNumber()
        {
            return number;
        }

        @Override public void setNumber(int n)
        {
            this.number = n;
        }
    }
}
