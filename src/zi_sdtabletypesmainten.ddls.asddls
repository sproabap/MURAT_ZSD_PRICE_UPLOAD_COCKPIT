@EndUserText.label: 'SD Table Types Maintenance'
@AccessControl.authorizationCheck: #CHECK
@Metadata.allowExtensions: true
define view entity ZI_SdTableTypesMainten
  as select from zsd_t_tabletypes
  association to parent ZI_SdTableTypesMainten_S as _SdTableTypesMainAll on $projection.SingletonID = _SdTableTypesMainAll.SingletonID
{
  key attachid as Attachid,
  key tabletype as Tabletype,
  @Semantics.largeObject: { mimeType: 'Mimetype',
                              fileName: 'Filename',
                              contentDispositionPreference: #ATTACHMENT,
                              acceptableMimeTypes: [ 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ]
                              }
  attachment as Attachment,
  @Semantics.mimeType: true
  @UI.hidden
  mimetype as Mimetype,
  filename as Filename,
  @Semantics.user.createdBy: true
  created_by as CreatedBy,
  @Semantics.systemDateTime.createdAt: true
  created_at as CreatedAt,
  @Semantics.user.lastChangedBy: true
  last_changed_by as LastChangedBy,
  @Semantics.systemDateTime.lastChangedAt: true
  last_changed_at as LastChangedAt,
  @Semantics.systemDateTime.localInstanceLastChangedAt: true
  @Consumption.hidden: true
  local_last_changed_at as LocalLastChangedAt,
  @Consumption.hidden: true
  1 as SingletonID,
  _SdTableTypesMainAll
  
}
