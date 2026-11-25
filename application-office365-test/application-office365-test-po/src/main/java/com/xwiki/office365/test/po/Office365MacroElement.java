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
package com.xwiki.office365.test.po;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.xwiki.test.ui.po.BaseElement;

/**
 * The output of an office365 macro. The tests wrap each macro in a group with an identifier, e.g.
 * {@code (% id="macro1" %)((({{office365/}})))}, to tell the macros of the same page apart.
 *
 * @version $Id$
 */
public class Office365MacroElement extends BaseElement
{
    private static final String VIEW_LINK = ".drive-links a[target='_blank']";

    private final String id;

    /**
     * @param id the identifier of the group containing the macro
     */
    public Office365MacroElement(String id)
    {
        this.id = id;
    }

    /**
     * @return the text displayed by the macro, including its error messages
     */
    public String getText()
    {
        return getContainer().getText();
    }

    /**
     * @return {@code true} if the macro displays a document in an iframe
     */
    public boolean hasIframe()
    {
        return has("iframe");
    }

    /**
     * @return the URL of the document displayed in the iframe
     */
    public String getIframeURL()
    {
        return getIframeAttribute("src");
    }

    /**
     * @return the width of the iframe
     */
    public String getIframeWidth()
    {
        return getIframeAttribute("width");
    }

    /**
     * @return the height of the iframe
     */
    public String getIframeHeight()
    {
        return getIframeAttribute("height");
    }

    /**
     * @return {@code true} if the macro failed, displaying an error instead of its output
     */
    public boolean hasError()
    {
        return has(".errormessage, .xwikirenderingerror");
    }

    /**
     * @return {@code true} if the macro displays the links of a synchronized document (change and view)
     */
    public boolean hasDocumentLinks()
    {
        return has(".drive-links");
    }

    /**
     * @return the URL of the "View" link, opening the synchronized document in Microsoft 365 in a new tab, or
     *     {@code null} if the link is not displayed
     */
    public String getViewURL()
    {
        return getContainer().findElements(By.cssSelector(VIEW_LINK)).stream().findFirst()
            .map(link -> link.getAttribute("href")).orElse(null);
    }

    /**
     * Click on the "Change" link, to replace the synchronized document, and wait for the page to reload.
     */
    public void clickChange()
    {
        getDriver().addPageNotYetReloadedMarker();
        getContainer().findElement(By.cssSelector(".drive-links a")).click();
        getDriver().waitUntilPageIsReloaded();
    }

    private WebElement getContainer()
    {
        return getDriver().findElement(By.id(this.id));
    }

    private boolean has(String cssSelector)
    {
        return !getContainer().findElements(By.cssSelector(cssSelector)).isEmpty();
    }

    private String getIframeAttribute(String name)
    {
        return getContainer().findElement(By.tagName("iframe")).getAttribute(name);
    }
}
