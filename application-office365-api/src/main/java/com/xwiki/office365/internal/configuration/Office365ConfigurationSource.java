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
package com.xwiki.office365.internal.configuration;

import javax.inject.Named;
import javax.inject.Singleton;

import org.xwiki.component.annotation.Component;
import org.xwiki.configuration.internal.AbstractWikisConfigurationSource;
import org.xwiki.model.reference.LocalDocumentReference;

/**
 * Reads the {@code Office365.Office365ConfigClass} object of the {@code Office365.Office365Config} page, from the
 * current wiki first, then from the main wiki.
 *
 * @version $Id$
 * @since 1.14.0
 */
@Component
@Named("office365")
@Singleton
public class Office365ConfigurationSource extends AbstractWikisConfigurationSource
{
    private static final String SPACE = "Office365";

    private static final LocalDocumentReference CONFIG_DOC = new LocalDocumentReference(SPACE, "Office365Config");

    private static final LocalDocumentReference CONFIG_CLASS =
        new LocalDocumentReference(SPACE, "Office365ConfigClass");

    @Override
    protected LocalDocumentReference getLocalDocumentReference()
    {
        return CONFIG_DOC;
    }

    @Override
    protected LocalDocumentReference getClassReference()
    {
        return CONFIG_CLASS;
    }

    @Override
    protected String getCacheId()
    {
        return "configuration.document.office365";
    }
}
