package microsoft.exchange.webservices.data.core.request;

import microsoft.exchange.webservices.data.ISelfValidate;
import microsoft.exchange.webservices.data.core.EwsServiceXmlWriter;
import microsoft.exchange.webservices.data.core.EwsUtilities;
import microsoft.exchange.webservices.data.core.XmlElementNames;
import microsoft.exchange.webservices.data.core.enumeration.misc.XmlNamespace;
import microsoft.exchange.webservices.data.property.complex.ComplexProperty;
import microsoft.exchange.webservices.data.property.complex.ConversationId;

public class ConversationRequest extends ComplexProperty implements ISelfValidate{
/// <summary>
  /// Gets or sets the conversation id.
  /// </summary>
  public ConversationId conversationId;

  public ConversationId getConversationId() {
    return conversationId;
  }

  public void setConversationId(ConversationId conversationId) {
    this.conversationId = conversationId;
  }

  public String getSyncState() {
    return syncState;
  }

  public void setSyncState(String syncState) {
    this.syncState = syncState;
  }

  /// <summary>
  /// Gets or sets the sync state representing the current state of the conversation for synchronization purposes.
  /// </summary>
  public String syncState;
  
  public ConversationRequest()
  {
  }

  /// <summary>
  /// Initializes a new instance of the <see cref="ConversationRequest"/> class.
  /// </summary>
  /// <param name="conversationId">The conversation id.</param>
  /// <param name="syncState">State of the sync.</param>
  public ConversationRequest(ConversationId conversationId, String syncState)
  {
      this.conversationId = conversationId;
      this.syncState = syncState;
  }
  
  /// <summary>
  /// Writes to XML.
  /// </summary>
  /// <param name="writer">The writer.</param>
  /// <param name="xmlElementName">Name of the XML element.</param>
  public void writeToXml(EwsServiceXmlWriter writer, String xmlElementName) throws Exception
  {
      writer.writeStartElement(XmlNamespace.Types, xmlElementName);

      this.conversationId.writeToXml(writer);

      if (this.syncState != null)
      {
          writer.writeElementValue(XmlNamespace.Types, XmlElementNames.SyncState, this.syncState);
      }

      writer.writeEndElement();
  }

  /// <summary>
  /// Validates this instance.
  /// </summary>
  protected void internalValidate() throws Exception
  {
      EwsUtilities.validateParam(this.conversationId, "ConversationId");
  }
  
}
