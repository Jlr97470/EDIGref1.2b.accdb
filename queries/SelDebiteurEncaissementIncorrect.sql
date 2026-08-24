SELECT
  DEBITEUR.Déb_Code,
  DEBITEUR.Déb_Situation,
  DEBITEUR.Déb_Date_Cré,
  DEBITEUR.Déb_Mont_Prin_m1,
  DEBITEUR.Déb_Mont_Prin_m2,
  DEBITEUR.Déb_MonnaieEnCours,
  (
    SELECT
      Round(
        Nz(
          Sum(Enc_Prin_m1),
          0
        ),
        2
      )
    FROM
      ENCAISSEMENT
    WHERE
      Enc_Code_Débiteur = Déb_Code;
  ) AS Déb_Enc_Prin_m1,
  (
    SELECT
      Round(
        Nz(
          Sum(Enc_Prin_m2),
          0
        ),
        2
      )
    FROM
      ENCAISSEMENT
    WHERE
      Enc_Code_Débiteur = Déb_Code;
  ) AS Déb_Enc_Prin_m2
FROM
  DEBITEUR
WHERE
  (
    (
      (DEBITEUR.Déb_Situation) = "99"
    )
    AND (
      (
        (
          SELECT
            Round(
              Nz(
                Sum(Enc_Prin_m1),
                0
              ),
              2
            )
          FROM
            ENCAISSEMENT
          WHERE
            Enc_Code_Débiteur = Déb_Code;
        )
      ) < Round([Déb_Mont_Prin_m1], 2)
    )
    AND (
      (
        (
          SELECT
            Round(
              Nz(
                Sum(Enc_Prin_m2),
                0
              ),
              2
            )
          FROM
            ENCAISSEMENT
          WHERE
            Enc_Code_Débiteur = Déb_Code;
        )
      ) < Round([Déb_Mont_Prin_m2], 2)
    )
  );
