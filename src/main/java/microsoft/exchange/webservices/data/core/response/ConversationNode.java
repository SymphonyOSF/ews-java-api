package microsoft.exchange.webservices.data.core.response;

import java.util.List;

import microsoft.exchange.webservices.data.core.EwsServiceXmlReader;
import microsoft.exchange.webservices.data.core.EwsUtilities;
import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.XmlElementNames;
import microsoft.exchange.webservices.data.core.service.ServiceObject;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.property.complex.ComplexProperty;
import microsoft.exchange.webservices.data.property.complex.ItemCollection;

public class ConversationNode extends ComplexProperty implements IGetObjectInstanceDelegate<ServiceObject> {

  private PropertySet propertySet;
  public String internetMessageId;
  public String parentInternetMessageId;
  public List<Item> items;
  private ItemCollection<Item> collections;

  public PropertySet getPropertySet() {
    return propertySet;
  }

  public void setPropertySet(PropertySet propertySet) {
    this.propertySet = propertySet;
  }

  public String getInternetMessageId() {
    return internetMessageId;
  }

  public void setInternetMessageId(String internetMessageId) {
    this.internetMessageId = internetMessageId;
  }

  public String getParentInternetMessageId() {
    return parentInternetMessageId;
  }

  public void setParentInternetMessageId(String parentInternetMessageId) {
    this.parentInternetMessageId = parentInternetMessageId;
  }

  public List<Item> getItems() {
    return items;
  }

  public void setItems(List<Item> items) {
    this.items = items;
  }

  public ConversationNode(PropertySet propertySet) {
    this.propertySet = propertySet;
  }

  @Override
  public boolean tryReadElementFromXml(EwsServiceXmlReader reader) throws Exception {
    if (reader.getLocalName().equals(XmlElementNames.InternetMessageId)) {
      this.internetMessageId = reader.readElementValue();
      return true;
    } else if (reader.getLocalName().equals(XmlElementNames.ParentInternetMessageId)) {
      this.parentInternetMessageId = reader.readElementValue();
      return true;
    } else if (reader.getLocalName().equals(XmlElementNames.Items)) {
      this.collections = new ItemCollection();
      this.collections.loadFromXml(reader, XmlElementNames.Items);
      this.items = this.collections.getItems();
      return true;
    } else {
      return false;
    }
  }

/// <summary>
  /// Gets the item instance.
  /// </summary>
  /// <param name="service">The service.</param>
  /// <param name="xmlElementName">Name of the XML element.</param>
  /// <returns>Item.</returns>
  @Override
  public ServiceObject getObjectInstanceDelegate(ExchangeService service, String xmlElementName) throws Exception {
    return this.getObjectInstance(service, xmlElementName);
  }

  private Folder getObjectInstance(ExchangeService service, String xmlElementName) throws Exception {
    return EwsUtilities.createEwsObjectFromXmlElementName(Item.class, service, xmlElementName);
  }

  /// <summary>
  /// Gets the name of the XML element.
  /// </summary>
  /// <returns>XML element name.</returns>
  public String getXmlElementName() {
    return XmlElementNames.ConversationNode;
  }
}
