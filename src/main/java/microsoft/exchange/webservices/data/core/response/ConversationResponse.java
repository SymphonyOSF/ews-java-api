package microsoft.exchange.webservices.data.core.response;

import microsoft.exchange.webservices.data.core.EwsServiceXmlReader;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.XmlElementNames;
import microsoft.exchange.webservices.data.property.complex.ComplexProperty;
import microsoft.exchange.webservices.data.property.complex.ConversationId;

public class ConversationResponse extends ComplexProperty {
/// <summary>
  /// Property set used to fetch items in the conversation.
  /// </summary>
  private PropertySet propertySet;
  
/// <summary>
  /// Gets the sync state.
  /// </summary>
  public String syncState;

  /// <summary>
  /// Gets the conversation nodes.
  /// </summary>
  public ConversationNodeCollection conversationNodes;

  /// <summary>
  /// Initializes a new instance of the <see cref="ConversationResponse"/> class.
  /// </summary>
  /// <param name="propertySet">The property set.</param>
  public ConversationResponse(PropertySet propertySet)
  {
      this.propertySet = propertySet;
  }

  /// <summary>
  /// Gets the conversation id.
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

  public ConversationNodeCollection getConversationNodes() {
    return conversationNodes;
  }

  public void setConversationNodes(ConversationNodeCollection conversationNodes) {
    this.conversationNodes = conversationNodes;
  }

  

  /// <summary>
  /// Tries to read element from XML.
  /// </summary>
  /// <param name="reader">The reader.</param>
  /// <returns>True if element was read.</returns>
  @Override
  public boolean tryReadElementFromXml(EwsServiceXmlReader reader) throws Exception
  {
    if(reader.getLocalName().equals(XmlElementNames.ConversationId)) {
      this.conversationId = new ConversationId();
      this.conversationId.loadFromXml(reader, XmlElementNames.ConversationId);
      return true;
    } else if(reader.getLocalName().equals(XmlElementNames.SyncState)) {
      this.syncState = reader.readElementValue();
      return true;
    } else if(reader.getLocalName().equals(XmlElementNames.ConversationNodes)) {
      this.conversationNodes = new ConversationNodeCollection(this.propertySet);
      this.conversationNodes.loadFromXml(reader, XmlElementNames.ConversationNodes);
      return true;
    } else {
      return false;
    }
  }
}
