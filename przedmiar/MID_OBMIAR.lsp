;;; MID_OBMIAR.lsp - eksport obmiaru z rysunku do pliku CSV (TYLKO ODCZYT RYSUNKU)
;;; Pracownia Projektowa MiD Sp. z o.o., 2026-10-26, wersja 1.0
;;;
;;; Polecenie: MID_OBMIAR
;;;   1. Pyta o filtr warstw (Enter = wszystkie, np. D_*,4*,5* albo *humus*).
;;;   2. Przechodzi po obiektach przestrzeni modelu biezacego rysunku
;;;      (bez zawartosci xrefow - uruchom w kazdym rysunku zrodlowym osobno).
;;;   3. Zapisuje obok rysunku plik RRRRMMDD_GGMMSS_<nazwa>_obmiar.csv
;;;      kolumny: uchwyt;warstwa;typ;blok;dlugosc_m;pole_m2;zamknieta;x;y
;;;      - dlugosc: krzywe (linie, polilinie, luki, okregi, splajny), obwod REGION
;;;      - pole: kreskowania (netto, z wyspami), REGION, okregi, polilinie (zamknieta=nie
;;;        oznacza pole liczone tak, jakby polilinia byla domknieta)
;;;   4. Wypisuje w oknie polecen sumy wg warstw.
;;; Rysunek nie jest zmieniany. Jednostki rysunku = metry (jak w planach MiD).

(vl-load-com)

(defun mid:try (f args / r)
  (setq r (vl-catch-all-apply f args))
  (if (vl-catch-all-error-p r) nil r)
)

(defun mid:num (x)
  (if (numberp x) (vl-string-translate "," "." (rtos x 2 4)) "")
)

(defun mid:csvtxt (s)
  ;; srednik w nazwie warstwy/bloku zamieniamy na przecinek, zeby nie rozbic kolumn
  (if s (vl-string-translate ";" "," s) "")
)

(defun mid:dlugosc (obj nazwa / ep)
  (cond
    ((= nazwa "AcDbRegion") (mid:try 'vla-get-Perimeter (list obj)))
    ((member nazwa '("AcDbLine" "AcDbPolyline" "AcDb2dPolyline" "AcDb3dPolyline"
                     "AcDbArc" "AcDbCircle" "AcDbEllipse" "AcDbSpline"))
      (setq ep (mid:try 'vlax-curve-getEndParam (list obj)))
      (if ep (mid:try 'vlax-curve-getDistAtParam (list obj ep)))
    )
    ;; obiekty Civil 3D (trasy, linie charakterystyczne) - proba przez interfejs krzywej
    ((wcmatch nazwa "Aecc*")
      (setq ep (mid:try 'vlax-curve-getEndParam (list obj)))
      (if ep (mid:try 'vlax-curve-getDistAtParam (list obj ep)))
    )
    (t nil)
  )
)

(defun mid:pole (obj nazwa)
  (if (member nazwa '("AcDbHatch" "AcDbRegion" "AcDbCircle" "AcDbEllipse"
                      "AcDbPolyline" "AcDb2dPolyline"))
    (mid:try 'vla-get-Area (list obj))
  )
)

(defun mid:zamknieta (obj nazwa / c)
  (cond
    ((member nazwa '("AcDbPolyline" "AcDb2dPolyline" "AcDb3dPolyline"))
      (setq c (mid:try 'vla-get-Closed (list obj)))
      (if (= c :vlax-true) "tak" "nie")
    )
    ((member nazwa '("AcDbCircle" "AcDbHatch" "AcDbRegion")) "tak")
    (t "")
  )
)

(defun mid:blok (obj nazwa / n)
  (if (= nazwa "AcDbBlockReference")
    (progn
      (setq n (mid:try 'vla-get-EffectiveName (list obj)))
      (if (not n) (setq n (mid:try 'vla-get-Name (list obj))))
      (if n n "")
    )
    ""
  )
)

(defun mid:srodek (obj / mn mx r)
  (setq r (vl-catch-all-apply 'vla-GetBoundingBox (list obj 'mn 'mx)))
  (if (and (not (vl-catch-all-error-p r)) mn mx)
    (progn
      (setq mn (vlax-safearray->list mn)
            mx (vlax-safearray->list mx))
      (list (/ (+ (car mn) (car mx)) 2.0) (/ (+ (cadr mn) (cadr mx)) 2.0))
    )
  )
)

(defun mid:dodaj (lst klucz dl pl / para)
  ;; lst: ((warstwa liczba dlugosc pole) ...)
  (if (setq para (assoc klucz lst))
    (subst (list klucz (1+ (cadr para))
                 (+ (caddr para) (if dl dl 0.0))
                 (+ (cadddr para) (if pl pl 0.0)))
           para lst)
    (cons (list klucz 1 (if dl dl 0.0) (if pl pl 0.0)) lst)
  )
)

(defun c:MID_OBMIAR (/ filtr doc ms plik f n obj nazwa warstwa dl pl sr sumy stamp)
  (setq filtr (getstring T "\nFiltr warstw <*>: "))
  (if (= filtr "") (setq filtr "*"))
  (setq doc (vla-get-ActiveDocument (vlax-get-acad-object))
        ms  (vla-get-ModelSpace doc)
        stamp (menucmd "M=$(edtime,$(getvar,date),YYYYMODD_HHMMSS)")
        plik (strcat (getvar "DWGPREFIX") stamp "_" (vl-filename-base (getvar "DWGNAME")) "_obmiar.csv")
        f (open plik "w")
        n 0
        sumy nil)
  (if (not f)
    (princ (strcat "\nNie mozna zapisac pliku: " plik))
    (progn
      (write-line "uchwyt;warstwa;typ;blok;dlugosc_m;pole_m2;zamknieta;x;y" f)
      (vlax-for obj ms
        (setq warstwa (vla-get-Layer obj)
              nazwa   (vla-get-ObjectName obj))
        (if (wcmatch (strcase warstwa) (strcase filtr))
          (progn
            (setq dl (mid:dlugosc obj nazwa)
                  pl (mid:pole obj nazwa)
                  sr (mid:srodek obj)
                  n  (1+ n)
                  sumy (mid:dodaj sumy warstwa dl pl))
            (write-line
              (strcat (vla-get-Handle obj) ";" (mid:csvtxt warstwa) ";" nazwa ";"
                      (mid:csvtxt (mid:blok obj nazwa)) ";" (mid:num dl) ";" (mid:num pl) ";"
                      (mid:zamknieta obj nazwa) ";"
                      (if sr (mid:num (car sr)) "") ";" (if sr (mid:num (cadr sr)) ""))
              f)
          )
        )
      )
      (close f)
      (princ "\n--- MID_OBMIAR: sumy wg warstw (liczba | dlugosc [m] | pole [m2]) ---")
      (foreach s (vl-sort sumy '(lambda (a b) (< (car a) (car b))))
        (princ (strcat "\n" (car s) " | " (itoa (cadr s)) " | "
                       (rtos (caddr s) 2 1) " | " (rtos (cadddr s) 2 1)))
      )
      (princ (strcat "\n--- Zapisano " (itoa n) " obiektow do: " plik))
    )
  )
  (princ)
)

(princ "\nMID_OBMIAR wczytany. Wpisz MID_OBMIAR, aby wyeksportowac obmiar do CSV.")
(princ)
