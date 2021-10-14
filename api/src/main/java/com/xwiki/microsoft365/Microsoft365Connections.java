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
package com.xwiki.microsoft365;

import java.util.List;

/**
 * Set of methods that the connections follow.
 *
 * @version $Id$
 * @since 1.0
 */
public interface Microsoft365Connections
{
    /**
     * Performs the search with queries in the request.
     *
     * @return a list {@link SearchResultItem}.
     */
    SearchResult searchDocuments();

    /**
     * Evaluates if the services cannot be offered before an authentication is performed with the Microsoft cloud.
     *
     * @return Returns true if the authentication is needed.
     */
    boolean isMissingAuth();

    /**
     * Returns an HTML string containing debugging information (the information can be presented to the user).
     *
     * @return An HTML string.
     */
    String getDebugInfo();

    /**
     * @return the URL to request to authorize.
     */
    String getOAuthStartUrl();

    /**
     * Performs the macro's preliminary tasks and makes the information available for the rendering of the macro to
     * happen.
     *
     * @param macroObject The XWiki macroObject
     * @return a set of information velocity rendering.
     */
    MacroRun runMacro(Object macroObject);

    /**
     * Retrieves the list of configured Sharepoint sites.
     *
     * @return The list of URLs of the sharepoint sites added to the configuration.
     */
    List<String> getSites();

    /**
     * Methods to drive the rendering of the macro.
     */
    interface MacroRun
    {
        /**
         * @return A string among displayPDF, displayEmbedIFrame, displaySearch,
         * displaySearchResults, and displayError.
         */
        String getMode();

        /**
         * @return true if the user lacks authentication to activate the services.
         */
        boolean isAuthenticationNeeded();

        /**
         * @return true if a redirect has been requested in case of lack of authentication.
         */
        boolean isRedirecting();

        /**
         * Used when mode is displayEmbedIFrame or displayPDF.
         *
         * @return the URL to the iframe.
         */
        String getUrl();

        /**
         * Used when mode is displayEmbedIFrame or displayPDF.
         *
         * @return the width of the frame
         */
        String getWidth();

        /**
         * Used when mode is displayEmbedIFrame or displayPDF.
         *
         * @return The height of the frame
         */
        String getHeight();

        /**
         * @return the title of an error. If null or empty, the run has been successful.
         */
        String getError();

        /**
         * @return the message of the error, useful if {@link #getError()} returns non-null.
         */
        String getErrorMessage();

        /**
         * @return the number parameter or, if unavailable, the sequence number.
         */
        int getNumber();

        /**
         * @param n the number assigned to the macro
         */
        void setNumber(int n);

    }

    /**
     * A search result.
     */
    interface SearchResultItem
    {
        /**
         * @return The name (title) of the document.
         */
        String getName();


        /**
         * @return The URL to trigger an embed of the document.
         */
        String getEmbedUrl();

        /**
         * @return The name to send the user to view the document.
         */
        String getViewUrl();

        /**
         * @return the cloud internal ID of the document.
         */
        String getId();

        /**
         * @return the current version number.
         */
        String getVersion();

        /**
         * @return the expected filename.
         */
        String getFilename();

    }

    /**
     * POJO to describe the result of a search.
     */
    class SearchResult
    {

        private List<SearchResultItem> items;
        private String error;
        private String message;
        private String searchedText;

        public SearchResult(List<SearchResultItem> items, String searchedText,  Exception ex) {
            this.items = items;
            if (ex != null) {
                this.error = ex.getMessage();
                Throwable cause = ex.getCause();
                if (cause != null) {
                    this.message = cause.getMessage();
                }
            }
            this.searchedText = searchedText;
            this.error = error;
            this.message = message;
        }

        /**
         * @return A non-null string if an error occurred.
         * In this case {@link #getErrorMessage()} should also return
         * a non-mepty string.
         */
        public String getError() {
            return error;
        }

        /**
         * @return The details of the error.
         */
        public String getErrorMessage() {
            return message;
        }

        public List<SearchResultItem> getItems() {
            return items;
        }

        /**
         * @return The text that was searched (as reported by the server).
         */
        public String getSearchedText() {
            return searchedText;
        }

    }
}
