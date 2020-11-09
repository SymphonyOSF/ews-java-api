package microsoft.exchange.webservices.data.core.response;

import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.property.complex.ComplexPropertyCollection;

public class ConversationNodeCollection extends ComplexPropertyCollection<ConversationNode>{

  private PropertySet propertySet;
  
  public ConversationNodeCollection(PropertySet propertySet) {
    this.propertySet = propertySet;
  }
  
  @Override
  protected ConversationNode createComplexProperty(String xmlElementName) {
    return new ConversationNode(this.propertySet);
  }

  @Override
  protected String getCollectionItemXmlElementName(ConversationNode complexProperty) {
    return complexProperty.getXmlElementName();
  }

}
