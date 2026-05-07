CLASS lhc_zr_sd_t_excel DEFINITION INHERITING FROM cl_abap_behavior_handler.
  PUBLIC SECTION.

  PRIVATE SECTION.
    METHODS:
      get_instance_features FOR INSTANCE FEATURES
        IMPORTING keys REQUEST requested_features FOR ZrSdTExcel RESULT result,
      get_global_authorizations FOR GLOBAL AUTHORIZATION
        IMPORTING
        REQUEST requested_authorizations FOR ZrSdTExcel
        RESULT result,
      updatetable FOR MODIFY
        IMPORTING keys   FOR ACTION ZrSdTExcel~updateTable
        RESULT    result.
ENDCLASS.

CLASS lhc_zr_sd_t_excel IMPLEMENTATION.
  METHOD get_global_authorizations.
  ENDMETHOD.

  METHOD updatetable.
    READ ENTITIES OF zr_sd_t_excel IN LOCAL MODE
         ENTITY ZrSdTExcel
         ALL FIELDS WITH CORRESPONDING #( keys )
         RESULT DATA(lt_result).

    IF lt_result IS INITIAL.
      RETURN.
    ENDIF.

    DATA(ls_result) = lt_result[ 1 ].
    DATA lt_response TYPE zcl_sd_types=>tt_response.

    IF ls_result-Attachment IS INITIAL.
      APPEND VALUE #( %tky = lt_result[ 1 ]-%tky
                      %msg = new_message( severity = if_abap_behv_message=>severity-error
                                          id       = 'ZSD_PRICE_MSG_CLS'
                                          number   = '004' ) ) TO reported-zrsdtexcel.
    ELSE.
      CASE ls_result-Tabletype.
        WHEN '01'.
          zcl_sd_excel=>read_shipping_excel( EXPORTING attachment  = ls_result-Attachment
                                             IMPORTING et_response = lt_response ).
        WHEN '02'.
          zcl_sd_excel=>read_km_excel( EXPORTING attachment  = ls_result-Attachment
                                       IMPORTING et_response = lt_response ).
        WHEN '03'.
          zcl_sd_excel=>read_pallet_excel( EXPORTING attachment  = ls_result-Attachment
                                           IMPORTING et_response = lt_response ).
        WHEN '04'.
          zcl_sd_excel=>read_weight_excel( EXPORTING attachment  = ls_result-Attachment
                                           IMPORTING et_response = lt_response ).
        WHEN '05'.
          zcl_sd_excel=>read_route_excel( EXPORTING attachment  = ls_result-Attachment
                                          IMPORTING et_response = lt_response ).
        WHEN '06'.
          zcl_sd_excel=>read_fixedrent_excel( EXPORTING attachment  = ls_result-Attachment
                                              IMPORTING et_response = lt_response ).
        WHEN '07'.
          zcl_sd_excel=>read_pointdiff_excel( EXPORTING attachment  = ls_result-Attachment
                                              IMPORTING et_response = lt_response ).
        WHEN '08'.
          zcl_sd_excel=>read_kmdiff_excel( EXPORTING attachment  = ls_result-Attachment
                                           IMPORTING et_response = lt_response ).
        WHEN '09'.
          zcl_sd_excel=>read_porterage_excel( EXPORTING attachment  = ls_result-Attachment
                                              IMPORTING et_response = lt_response ).
        WHEN '10'.
          zcl_sd_excel=>read_waiting_excel( EXPORTING attachment  = ls_result-Attachment
                                            IMPORTING et_response = lt_response ).
        WHEN '11'.
          zcl_sd_excel=>read_tollcost_excel( EXPORTING attachment  = ls_result-Attachment
                                             IMPORTING et_response = lt_response ).
      ENDCASE.

      LOOP AT lt_response INTO DATA(ls_response).
        CASE ls_response-flag.
          WHEN '1'.
            APPEND VALUE #( %tky = lt_result[ 1 ]-%tky
                            %msg = new_message( id       = 'ZSD_PRICE_MSG_CLS'
                                                number   = 001
                                                severity = if_abap_behv_message=>severity-success ) )
                   TO reported-zrsdtexcel.
          WHEN '2'.
            APPEND VALUE #( %tky = lt_result[ 1 ]-%tky ) TO failed-zrsdtexcel.
            APPEND VALUE #( %tky = lt_result[ 1 ]-%tky
                            %msg = new_message( id       = 'ZSD_PRICE_MSG_CLS'
                                                number   = 002
                                                severity = if_abap_behv_message=>severity-error
                                                v1       = ls_response-tabix ) )
                   TO reported-zrsdtexcel.
          WHEN '3'.
            APPEND VALUE #( %tky = lt_result[ 1 ]-%tky ) TO failed-zrsdtexcel.
            APPEND VALUE #( %tky = lt_result[ 1 ]-%tky
                            %msg = new_message( id       = 'ZSD_PRICE_MSG_CLS'
                                                number   = 003
                                                severity = if_abap_behv_message=>severity-error
                                                v1       = ls_response-tabix
                                                v2       = ls_response-field ) )
                   TO reported-zrsdtexcel.
          WHEN '4'.
            APPEND VALUE #( %tky = lt_result[ 1 ]-%tky ) TO failed-zrsdtexcel.
            APPEND VALUE #( %tky = lt_result[ 1 ]-%tky
                            %msg = new_message( id       = 'ZSD_PRICE_MSG_CLS'
                                                number   = 005
                                                severity = if_abap_behv_message=>severity-error
                                                v1       = ls_response-tabix ) )
                   TO reported-zrsdtexcel.
          WHEN '5'.
            APPEND VALUE #( %tky = lt_result[ 1 ]-%tky ) TO failed-zrsdtexcel.
            APPEND VALUE #( %tky = lt_result[ 1 ]-%tky
                            %msg = new_message( id       = 'ZSD_PRICE_MSG_CLS'
                                                number   = 006
                                                severity = if_abap_behv_message=>severity-error
                                                v1       = ls_response-tabix ) )
                   TO reported-zrsdtexcel.
          WHEN '6'.
            APPEND VALUE #( %tky = lt_result[ 1 ]-%tky ) TO failed-zrsdtexcel.
            APPEND VALUE #( %tky = lt_result[ 1 ]-%tky
                            %msg = new_message( id       = 'ZSD_PRICE_MSG_CLS'
                                                number   = 007
                                                severity = if_abap_behv_message=>severity-error
                                                v1       = ls_response-tabix ) )
                   TO reported-zrsdtexcel.
          WHEN '7'.
            APPEND VALUE #( %tky = lt_result[ 1 ]-%tky ) TO failed-zrsdtexcel.
            APPEND VALUE #( %tky = lt_result[ 1 ]-%tky
                            %msg = new_message( id       = 'ZSD_PRICE_MSG_CLS'
                                                number   = 008
                                                severity = if_abap_behv_message=>severity-error ) )
                   TO reported-zrsdtexcel.
          WHEN '8'.
            APPEND VALUE #( %tky = lt_result[ 1 ]-%tky ) TO failed-zrsdtexcel.
            APPEND VALUE #( %tky = lt_result[ 1 ]-%tky
                            %msg = new_message( id       = 'ZSD_PRICE_MSG_CLS'
                                                number   = 009
                                                severity = if_abap_behv_message=>severity-error ) )
                   TO reported-zrsdtexcel.
        ENDCASE.
      ENDLOOP.
      result = VALUE #( FOR ls_resultt IN lt_result
                        ( %tky   = ls_resultt-%tky
                          %param = ls_resultt ) ).
    ENDIF.
  ENDMETHOD.

  METHOD get_instance_features.
    LOOP AT keys INTO DATA(key).

      APPEND VALUE #( %tky                = key-%tky
                      %action-updateTable = COND #( WHEN key-%is_draft = if_abap_behv=>mk-on
                                                    THEN if_abap_behv=>fc-o-disabled
                                                    ELSE if_abap_behv=>fc-o-enabled ) )
             TO result.
    ENDLOOP.
  ENDMETHOD.

ENDCLASS.

CLASS lsc_zr_sd_t_excel DEFINITION INHERITING FROM cl_abap_behavior_saver.
  PROTECTED SECTION.
ENDCLASS.

CLASS lsc_zr_sd_t_excel IMPLEMENTATION.
ENDCLASS.
