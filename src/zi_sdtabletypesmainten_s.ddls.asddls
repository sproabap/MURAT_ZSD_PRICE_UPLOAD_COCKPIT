@EndUserText.label: 'SD Table Types Maintenance Singleton'
@AccessControl.authorizationCheck: #NOT_REQUIRED
@ObjectModel.semanticKey: [ 'SingletonID' ]
@UI: {
  headerInfo: {
    typeName: 'SdTableTypesMainAll'
  }
}
define root view entity ZI_SdTableTypesMainten_S
  as select from I_Language
    left outer join ZSD_T_TABLETYPES on 0 = 0
  association [0..*] to I_ABAPTransportRequestText as _I_ABAPTransportRequestText on $projection.TransportRequestID = _I_ABAPTransportRequestText.TransportRequestID
  composition [0..*] of ZI_SdTableTypesMainten as _SdTableTypesMainten
{
  @UI.facet: [ {
    id: 'ZI_SdTableTypesMainten', 
    purpose: #STANDARD, 
    type: #LINEITEM_REFERENCE, 
    label: 'SD Table Types Maintenance', 
    position: 1 , 
    targetElement: '_SdTableTypesMainten'
  } ]
  @UI.lineItem: [ {
    position: 1 
  } ]
  key 1 as SingletonID,
  _SdTableTypesMainten,
  @UI.hidden: true
  max( ZSD_T_TABLETYPES.LAST_CHANGED_AT ) as LastChangedAtMax,
  @ObjectModel.text.association: '_I_ABAPTransportRequestText'
  @UI.identification: [ {
    position: 2 , 
    type: #WITH_INTENT_BASED_NAVIGATION, 
    semanticObjectAction: 'manage'
  } ]
  @Consumption.semanticObject: 'CustomizingTransport'
  cast( '' as SXCO_TRANSPORT) as TransportRequestID,
  _I_ABAPTransportRequestText
  
}
where I_Language.Language = $session.system_language
