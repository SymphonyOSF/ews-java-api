/*
 * The MIT License
 * Copyright (c) 2012 Microsoft Corporation
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

package microsoft.exchange.webservices.data.core.request;

import java.util.List;

import microsoft.exchange.webservices.data.core.EwsServiceXmlWriter;
import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.XmlElementNames;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.misc.XmlNamespace;
import microsoft.exchange.webservices.data.core.enumeration.service.ServiceObjectType;
import microsoft.exchange.webservices.data.core.enumeration.service.error.ServiceErrorHandling;
import microsoft.exchange.webservices.data.core.exception.service.local.ServiceVersionException;
import microsoft.exchange.webservices.data.core.response.GetConversationItemsResponse;
import microsoft.exchange.webservices.data.property.complex.FolderIdCollection;
import microsoft.exchange.webservices.data.search.MailboxSearchLocation;

/**
 * Represents an abstract GetItem request.
 */
public final class GetConversationItemsRequest extends MultiResponseServiceRequest<GetConversationItemsResponse> {

/// <summary>
  /// Initializes a new instance of the <see cref="GetConversationItemsRequest"/> class.
  /// </summary>
  /// <param name="service">The service.</param>
  /// <param name="errorHandlingMode">Error handling mode.</param>
  public GetConversationItemsRequest(ExchangeService service, ServiceErrorHandling errorHandlingMode) throws Exception
  {
    super(service, errorHandlingMode);
  }

  /// <summary>
  /// Gets or sets the conversations.
  /// </summary>
  private List<ConversationRequest> conversations;
  /// <summary>
  /// Gets or sets the item properties.
  /// </summary>
  private PropertySet itemProperties;

  /// <summary>
  /// Gets or sets the folders to ignore.
  /// </summary>
  private FolderIdCollection foldersToIgnore;

  /// <summary>
  /// Gets or sets the maximum number of items to return.
  /// </summary>
  private int maxItemsToReturn;

  private ConversationSortOrder sortOrder;

  /// <summary>
  /// Gets or sets the mailbox search location to include in the search.
  /// </summary>
  private MailboxSearchLocation mailboxScope;

  public List<ConversationRequest> getConversations() {
    return conversations;
  }
  public void setConversations(List<ConversationRequest> conversations) {
    this.conversations = conversations;
  }
  public PropertySet getItemProperties() {
    return itemProperties;
  }
  public void setItemProperties(PropertySet itemProperties) {
    this.itemProperties = itemProperties;
  }
  public FolderIdCollection getFoldersToIgnore() {
    return foldersToIgnore;
  }
  public void setFoldersToIgnore(FolderIdCollection foldersToIgnore) {
    this.foldersToIgnore = foldersToIgnore;
  }
  public int getMaxItemsToReturn() {
    return maxItemsToReturn;
  }
  public void setMaxItemsToReturn(int maxItemsToReturn) {
    this.maxItemsToReturn = maxItemsToReturn;
  }
  public ConversationSortOrder getSortOrder() {
    return sortOrder;
  }
  public void setSortOrder(ConversationSortOrder sortOrder) {
    this.sortOrder = sortOrder;
  }
  public MailboxSearchLocation getMailboxScope() {
    return mailboxScope;
  }
  public void setMailboxScope(MailboxSearchLocation mailboxScope) {
    this.mailboxScope = mailboxScope;
  }

  @Override
  protected String getResponseMessageXmlElementName() {
    
    return XmlElementNames.GetConversationItemsResponseMessage;
  }
  @Override
  protected int getExpectedResponseMessageCount() {
    
    return this.conversations.size();
  }
  @Override
  public String getXmlElementName() {
    
    return XmlElementNames.GetConversationItems;
  }
  @Override
  protected String getResponseXmlElementName() {
    
    return XmlElementNames.GetConversationItemsResponse;
  }
  @Override
  protected ExchangeVersion getMinimumRequiredServerVersion() {
    
    return ExchangeVersion.Exchange2013;
  }
  
  @Override
  protected void writeElementsToXml(EwsServiceXmlWriter writer) throws Exception {
    
    this.itemProperties.writeToXml(writer, ServiceObjectType.Item);
    
    if (foldersToIgnore != null) {
      this.foldersToIgnore.writeToXml(writer, XmlNamespace.Messages, XmlElementNames.FoldersToIgnore);
    }
    
    if (this.maxItemsToReturn > 0)
    {
        writer.writeElementValue(XmlNamespace.Messages, XmlElementNames.MaxItemsToReturn, this.maxItemsToReturn);
    }

    if (this.sortOrder != null)
    {
        writer.writeElementValue(XmlNamespace.Messages, XmlElementNames.SortOrder, this.sortOrder);
    }

    if (this.mailboxScope != null)
    {
        writer.writeElementValue(XmlNamespace.Messages, XmlElementNames.MailboxScope, this.mailboxScope);
    }

    writer.writeStartElement(XmlNamespace.Messages, XmlElementNames.Conversations);
    for(ConversationRequest conversation : this.conversations) {
      conversation.writeToXml(writer, XmlElementNames.Conversation);
    }
    writer.writeEndElement();
  }
  
  @Override
  protected GetConversationItemsResponse createServiceResponse(ExchangeService service, int responseIndex)
      throws Exception {
    return new GetConversationItemsResponse(this.itemProperties);
  }
  
  @Override
  protected void validate() throws Exception {
    super.validate();

    // SearchScope is only valid for Exchange2013 or higher
    //
    if (this.mailboxScope != null &&
        this.getService().getRequestedServerVersion().ordinal() < ExchangeVersion.Exchange2013.ordinal()) {
        throw new ServiceVersionException(
            String.format(
                "invalid",
                "MailboxScope",
                ExchangeVersion.Exchange2013));
    }
  }
}
