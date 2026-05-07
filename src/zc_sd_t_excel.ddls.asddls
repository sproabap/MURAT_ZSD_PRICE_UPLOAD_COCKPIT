@Metadata.allowExtensions: true
@EndUserText.label: 'Price Upload Cockpit'
@AccessControl.authorizationCheck: #CHECK
define root view entity ZC_SD_T_EXCEL
  provider contract transactional_query
  as projection on ZR_SD_T_EXCEL
{
  key Attachid,
      @ObjectModel.text.element: [ 'text' ]
  key Tabletype,
      Attachment,
      Mimetype,
      Filename,
      @ObjectModel.text.element: [ 'createdbytext' ]
      @UI.textArrangement: #TEXT_ONLY
      CreatedBy,
      CreatedAt,
      @ObjectModel.text.element: [ 'changedbytext' ]
      LastChangedBy,
      LastChangedAt,
      LocalLastChangedAt,

      _Table.text                as text,
      _UserCreate.PersonFullName as createdbytext,
      _UserChange.PersonFullName as changedbytext,
      
      url

}
