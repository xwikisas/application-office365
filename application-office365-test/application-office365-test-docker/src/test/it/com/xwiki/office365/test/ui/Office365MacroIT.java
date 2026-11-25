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
package com.xwiki.office365.test.ui;

import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.Order;
import org.junit.jupiter.api.Test;
import org.xwiki.model.reference.DocumentReference;
import org.xwiki.test.docker.junit5.UITest;
import org.xwiki.test.ui.TestUtils;

import com.xwiki.office365.test.po.Office365MacroElement;
import com.xwiki.office365.test.po.Office365TestUtils;

import static org.hamcrest.MatcherAssert.assertThat;
import static org.hamcrest.Matchers.containsString;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Tests the office365 macro, embedding a document either by URL or synchronized with Microsoft 365 (stored in an
 * Office365.SynchronizedDocumentClass object of the page). The Microsoft 365 access (tenant, client id and secret) is
 * not configured: embedding by URL doesn't need it, and a synchronized document is chosen through the embed link of
 * a search result, which is opened directly since searching requires a Microsoft 365 account.
 *
 * @version $Id$
 * @since 1.14.0
 */
@UITest(
    properties = {
        // The Office365 pages use Groovy, which requires programming rights.
        "xwikiPropertiesAdditionalProperties=test.prchecker.excludePattern=.*:Office365\\..*"
    }
)
class Office365MacroIT
{
    private static final String SPACE = "Office365Test";

    private static final String PDF_URL = "https://example.com/Document.pdf";

    private static final String EDIT_URL = "https://example.com/Document.docx?action=default";

    private static final String EMBED_URL = "https://example.com/Document.docx?action=embedview";

    private static final String OTHER_EMBED_URL = "https://example.com/Other.docx?action=embedview";

    private static final String INVALID_URL = "Invalid url";

    @BeforeAll
    static void beforeAll(TestUtils setup)
    {
        setup.loginAsSuperAdmin();
        // The macro without URL authenticates the user on Microsoft 365, which redirects to the Microsoft login page,
        // unless the authentication already failed too many times. Make it fail, so that the tests don't need
        // Microsoft 365: they also check that the macro keeps working when the authentication fails.
        Office365TestUtils.failAuthenticationTooManyTimes(setup);
    }

    @Test
    @Order(1)
    void embedByURL(TestUtils setup)
    {
        DocumentReference page = new DocumentReference("xwiki", List.of(SPACE), "EmbedByURL");
        setup.createPage(page, String.join("\n\n",
            Office365TestUtils.macro("default", "url=\"https://example.com/Document.pptx?action=default\""),
            Office365TestUtils.macro("size",
                "url=\"https://example.com/Document.docx\" width=\"50%\" height=\"300px\""),
            Office365TestUtils.macro("pdf", "url=\"" + PDF_URL + "\""),
            Office365TestUtils.macro("javascript", "url=\"javascript:alert(1)\""),
            Office365TestUtils.macro("ftp", "url=\"ftp://example.com/Document.docx\""),
            Office365TestUtils.macro("nohost", "url=\"https:///Document.docx\""),
            Office365TestUtils.macro("noprotocol", "url=\"example.com/Document.docx\""),
            Office365TestUtils.macro("width", "url=\"https://example.com/Document.docx\" width=\"abc\"")),
            "Embed by URL");

        // The document is displayed with the embed view of Microsoft 365.
        Office365MacroElement macro = new Office365MacroElement("default");
        assertEquals("https://example.com/Document.pptx?action=embedview", macro.getIframeURL());
        assertEquals("100%", macro.getIframeWidth());
        assertEquals("800", macro.getIframeHeight());

        macro = new Office365MacroElement("size");
        assertEquals("https://example.com/Document.docx", macro.getIframeURL());
        assertEquals("50%", macro.getIframeWidth());
        assertEquals("300px", macro.getIframeHeight());

        // Microsoft 365 doesn't allow embedding its PDF viewer.
        macro = new Office365MacroElement("pdf");
        assertFalse(macro.hasIframe());
        assertThat(macro.getText(), containsString("This macro cannot display a pdf document"));
        assertThat(macro.getText(), containsString(PDF_URL));

        for (String id : List.of("javascript", "ftp", "nohost", "noprotocol")) {
            macro = new Office365MacroElement(id);
            assertFalse(macro.hasIframe(), id);
            assertThat(macro.getText(), containsString(INVALID_URL));
        }
        macro = new Office365MacroElement("width");
        assertFalse(macro.hasIframe());
        assertThat(macro.getText(), containsString("Invalid width or height parameters."));

        // The PDF can still be displayed in an iframe when forced.
        setup.gotoPage(page, "view", "force=1");
        assertEquals(PDF_URL, new Office365MacroElement("pdf").getIframeURL());
        // But invalid urls still don't display.
        for (String id : List.of("javascript", "ftp", "nohost", "noprotocol")) {
            macro = new Office365MacroElement(id);
            assertFalse(macro.hasIframe(), id);
            assertThat(macro.getText(), containsString(INVALID_URL));
        }
    }

    @Test
    @Order(2)
    void embedAndChange(TestUtils setup)
    {
        DocumentReference page = new DocumentReference("xwiki", List.of(SPACE), "EmbedAndChange");
        setup.createPage(page, Office365TestUtils.macro("macro", ""), "Embed and Change");
        Office365MacroElement macro = new Office365MacroElement("macro");
        assertFalse(macro.hasIframe());
        assertFalse(macro.hasError());

        embed(setup, page, 0, EMBED_URL, EDIT_URL);
        assertSynchronizedDocument(macro, EMBED_URL, EDIT_URL);

        // The document is stored in the page.
        setup.gotoPage(page, "view", "reload=1");
        assertSynchronizedDocument(macro, EMBED_URL, EDIT_URL);

        // Changing the document shows the search instead of the document, without forgetting it.
        macro.clickChange();
        assertFalse(macro.hasIframe());
        assertFalse(macro.hasError());
        setup.gotoPage(page, "view", "reload=2");
        assertSynchronizedDocument(macro, EMBED_URL, EDIT_URL);

        // Embed another document. Its edit link is not an HTTP URL, so the "View" link is not displayed.
        embed(setup, page, 0, OTHER_EMBED_URL, "javascript:alert(1)");
        assertSynchronizedDocument(macro, OTHER_EMBED_URL, null);
    }

    @Test
    @Order(3)
    void embedInMacroWithNumber(TestUtils setup)
    {
        DocumentReference page = new DocumentReference("xwiki", List.of(SPACE), "EmbedInMacroWithNumber");
        setup.createPage(page,
            Office365TestUtils.macro("first", "nb=\"0\"") + "\n\n" + Office365TestUtils.macro("second", "nb=\"1\""),
            "Embed in Macro with Number");

        embed(setup, page, 1, EMBED_URL, EDIT_URL);
        assertFalse(new Office365MacroElement("first").hasIframe());
        assertSynchronizedDocument(new Office365MacroElement("second"), EMBED_URL, EDIT_URL);
    }

    @Test
    @Order(4)
    @Disabled("All the office365 macros of a page without the nb parameter share the same synchronized document")
    void embedInMacroWithoutNumber(TestUtils setup)
    {
        DocumentReference page = new DocumentReference("xwiki", List.of(SPACE), "EmbedInMacroWithoutNumber");
        setup.createPage(page, Office365TestUtils.macro("first", "") + "\n\n" + Office365TestUtils.macro("second", ""),
            "Embed in Macro without Number");

        embed(setup, page, 0, EMBED_URL, EDIT_URL);
        assertSynchronizedDocument(new Office365MacroElement("first"), EMBED_URL, EDIT_URL);
        assertFalse(new Office365MacroElement("second").hasIframe());
    }

    /**
     * Open the link to embed a search result, as displayed by the macro.
     */
    private void embed(TestUtils setup, DocumentReference page, int nb, String embedLink, String editLink)
    {
        Map<String, Object> parameters = new LinkedHashMap<>();
        parameters.put("embed", "1");
        parameters.put("nb", nb);
        parameters.put("id", "fake-item-id");
        parameters.put("filename", "Document.docx");
        parameters.put("embedLink", embedLink);
        parameters.put("editLink", editLink);
        parameters.put("site", "");
        parameters.put("form_token", setup.getSecretToken());
        setup.gotoPage(page, "view", parameters);
    }

    /**
     * @param embedURL the expected URL of the document displayed in the iframe
     * @param viewURL the expected URL of the "View" link, or {@code null} if it shouldn't be displayed
     */
    private void assertSynchronizedDocument(Office365MacroElement macro, String embedURL, String viewURL)
    {
        assertTrue(macro.hasDocumentLinks());
        assertEquals(embedURL, macro.getIframeURL());
        assertEquals(viewURL, macro.getViewURL());
    }
}
