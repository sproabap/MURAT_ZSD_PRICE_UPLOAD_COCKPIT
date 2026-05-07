@AbapCatalog.viewEnhancementCategory: [#NONE]
@AccessControl.authorizationCheck: #NOT_REQUIRED
@EndUserText.label: 'Table Types Value Help'
@Metadata.ignorePropagatedAnnotations: true
@ObjectModel.usageType:{
    serviceQuality: #X,
    sizeCategory: #S,
    dataClass: #MIXED
}
@ObjectModel.resultSet.sizeCategory: #XS
define view entity ZSD_I_TABLETYPES_VH as select from DDCDS_CUSTOMER_DOMAIN_VALUE_T( p_domain_name : 'ZSD_TABLETYPE_DO' )
{
    @UI.hidden: true
    key domain_name,
    @UI.hidden: true
    key value_position,
    @Semantics.language:true
    @UI.hidden: true
    key language,
    @UI.hidden: true
    value_low,
    @Semantics.text:true
    text
}where language = $session.system_language
