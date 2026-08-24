SELECT
  DEBITEUR.*,
  ADRESSE_DEBITEUR.*
FROM
  DEBITEUR
  INNER JOIN ADRESSE_DEBITEUR ON DEBITEUR.Déb_Code = ADRESSE_DEBITEUR.AD_Code_Débiteur
WHERE
  (
    (
      (
        ADRESSE_DEBITEUR.AD_TypeAdresse_ID
      ) = DLookUp(
        "Adresse_Courrier", "PARAMETRE_GREF"
      )
    )
  );
