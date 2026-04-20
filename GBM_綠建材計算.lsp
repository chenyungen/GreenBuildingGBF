;;; ============================================================
;;;  GBM — 綠建材空間計算工具  (for GreenBuildingGBF.exe)
;;;  陳國崇建築師事務所 × Claude · 2026
;;; ============================================================
;;;  指令
;;;    GBM            顯示指令選單
;;;    GBM-ADD        從封閉線自動擷取面積／周長，新增一個空間
;;;    GBM-HATCH      由剖面線 (HATCH) 統計綠建材面積
;;;    GBM-LIST       列出所有已記錄空間
;;;    GBM-EDIT       修改特定空間的綠建材面積
;;;    GBM-DELETE     刪除指定空間
;;;    GBM-CLEAR      清空所有資料
;;;    GBM-EXPORT     匯出為 CSV (給 GreenBuildingGBF.exe 讀取)
;;;    GBM-TABLE      在圖面上繪製計算表格
;;;    GBM-LOAD       從現有 CSV 載回資料
;;; ============================================================

(vl-load-com)

;; Global data: list of (id Ai1 Li k H2 Ai2 Gi1 Gi4 Mat1 Mat4 Ai4)
(if (not *GBM-DATA*) (setq *GBM-DATA* '()))
(if (not *GBM-H2*)   (setq *GBM-H2*   2.8))  ; default ceiling height
(if (not *GBM-K*)    (setq *GBM-K*    0.8))  ; default k
(if (not *GBM-PATTERN-MAP*)
  (setq *GBM-PATTERN-MAP* '(
    ("ANSI31" . "乳膠漆")
    ("AR-HBONE" . "矽酸鈣天花板")
    ("NET" . "矽酸鈣天花板")
    ("GRASS" . "乳膠漆")
  )))

;;; ------------------ Helpers ------------------

(defun gbm:prop (ename prop / obj r)
  (setq obj (vlax-ename->vla-object ename))
  (setq r (vl-catch-all-apply 'vlax-get-property (list obj prop)))
  (if (vl-catch-all-error-p r) nil r))

(defun gbm:area (e)
  (cond
    ((gbm:prop e 'Area))
    (T nil)))

(defun gbm:length (e / v)
  (cond
    ((setq v (gbm:prop e 'Length)) v)
    ((setq v (gbm:prop e 'Perimeter)) v)
    ((setq v (gbm:prop e 'Circumference)) v)
    (T nil)))

(defun gbm:fmt (n d) (rtos n 2 d))

(defun gbm:find-record (id / rec found)
  (setq found nil)
  (foreach rec *GBM-DATA*
    (if (equal (nth 0 rec) id) (setq found rec)))
  found)

(defun gbm:replace-record (id new-rec)
  (setq *GBM-DATA*
    (mapcar '(lambda (r) (if (equal (nth 0 r) id) new-rec r))
            *GBM-DATA*)))

;;; ------------------ GBM-ADD ------------------

(defun c:GBM-ADD ( / sel ent ai1 li id h2 k ai4 ai2 g1_pct g4_pct g1 g4 m1 m4 rec ai4_input)
  (princ "\n=== GBM-ADD 新增空間 ===")
  (setq sel (entsel "\n選擇房間邊界（封閉聚合線／圓／區域）: "))
  (cond
    ((null sel) (princ "\n取消。"))
    (T
      (setq ent (car sel))
      (setq ai1 (gbm:area ent))
      (setq li  (gbm:length ent))
      (cond
        ((or (null ai1) (< ai1 0.01))
          (princ "\n選取物件非封閉區域或面積為零。"))
        (T
          (princ (strcat "\n  天花板面積 Ai,1 = " (gbm:fmt ai1 2) " m²"))
          (princ (strcat "\n  內部周長   Li   = " (gbm:fmt li 2) " m"))
          (setq id (getstring (strcat "\n空間編號 <" (itoa (1+ (length *GBM-DATA*))) ">: ")))
          (if (= id "") (setq id (itoa (1+ (length *GBM-DATA*)))))
          (setq h2 (getreal (strcat "\n天花板高度 H2 (m) <" (gbm:fmt *GBM-H2* 2) ">: ")))
          (if (or (null h2) (<= h2 0)) (setq h2 *GBM-H2*))
          (setq *GBM-H2* h2)
          (setq k (getreal (strcat "\n係數 k <" (gbm:fmt *GBM-K* 2) ">: ")))
          (if (or (null k) (<= k 0)) (setq k *GBM-K*))
          (setq *GBM-K* k)
          (setq ai4_input (getreal (strcat "\n牆面面積 Ai,4 (m²) <" (gbm:fmt ai1 2) ">: ")))
          (if (or (null ai4_input) (<= ai4_input 0)) (setq ai4 ai1) (setq ai4 ai4_input))
          (setq ai2 (* ai4 k h2))
          (princ (strcat "\n  內部裝修總面積 Ai,2 = " (gbm:fmt ai4 2) "×" (gbm:fmt k 2) "×" (gbm:fmt h2 2) " = " (gbm:fmt ai2 2) " m²"))
          (setq g1_pct (getreal "\n天花板使用綠建材 % <0>: "))
          (if (or (null g1_pct) (< g1_pct 0)) (setq g1_pct 0))
          (if (> g1_pct 100) (setq g1_pct 100))
          (setq g1 (* ai1 (/ g1_pct 100.0)))
          (setq g4_pct (getreal "\n牆面使用綠建材 % <0>: "))
          (if (or (null g4_pct) (< g4_pct 0)) (setq g4_pct 0))
          (if (> g4_pct 100) (setq g4_pct 100))
          (setq g4 (* ai2 (/ g4_pct 100.0)))
          (setq m1 (getstring T "\n天花板材料 <矽酸鈣天花板>: "))
          (if (= m1 "") (setq m1 "矽酸鈣天花板"))
          (setq m4 (getstring T "\n牆面材料 <乳膠漆>: "))
          (if (= m4 "") (setq m4 "乳膠漆"))
          (setq rec (list id ai1 li k h2 ai2 g1 g4 m1 m4 ai4))
          (setq *GBM-DATA* (append *GBM-DATA* (list rec)))
          (princ (strcat "\n✔ 已新增空間 " id "。(目前共 " (itoa (length *GBM-DATA*)) " 筆)"))))))
  (princ))

;;; ------------------ GBM-HATCH ------------------
;;; Select multiple hatches, group by pattern, show total area per pattern.

(defun c:GBM-HATCH ( / ss i ent obj pat ar key found data)
  (princ "\n=== GBM-HATCH 剖面線面積統計 ===")
  (princ "\n選擇剖面線 (HATCH) 物件 (可多選):")
  (setq ss (ssget '((0 . "HATCH"))))
  (cond
    ((null ss) (princ "\n沒有選取物件。"))
    (T
      (setq data '() i 0)
      (while (< i (sslength ss))
        (setq ent (ssname ss i))
        (setq obj (vlax-ename->vla-object ent))
        (setq pat (vl-catch-all-apply 'vlax-get-property (list obj 'PatternName)))
        (setq ar  (vl-catch-all-apply 'vlax-get-property (list obj 'Area)))
        (if (and (not (vl-catch-all-error-p pat))
                 (not (vl-catch-all-error-p ar))
                 (numberp ar))
          (progn
            (setq key pat)
            (setq found (assoc key data))
            (if found
              (setq data (subst (cons key (+ (cdr found) ar)) found data))
              (setq data (append data (list (cons key ar)))))))
        (setq i (1+ i)))
      (princ "\n\n  剖面圖案          總面積 (m²)   對應材料")
      (princ "\n  ─────────────────────────────────────────")
      (foreach p data
        (princ (strcat "\n  "
          (substr (strcat (car p) "                ") 1 16)
          (gbm:fmt (cdr p) 3)
          "       "
          (cond ((cdr (assoc (car p) *GBM-PATTERN-MAP*)))
                (T "（未定義，見 *GBM-PATTERN-MAP*）")))))
      (princ "\n")))
  (princ))

;;; ------------------ GBM-LIST ------------------

(defun c:GBM-LIST ( / rec g1sum g4sum a1sum a2sum)
  (cond
    ((null *GBM-DATA*) (princ "\n（尚無資料）"))
    (T
      (setq g1sum 0 g4sum 0 a1sum 0 a2sum 0)
      (princ "\n=============== 已記錄空間 ===============")
      (princ "\n ID     Ai,1    Li     k   H2    Ai,2     Gi1     Gi4    Mat1/Mat4")
      (foreach rec *GBM-DATA*
        (setq a1sum (+ a1sum (nth 1 rec)))
        (setq a2sum (+ a2sum (nth 5 rec)))
        (setq g1sum (+ g1sum (nth 6 rec)))
        (setq g4sum (+ g4sum (nth 7 rec)))
        (princ (strcat "\n "
          (substr (strcat (nth 0 rec) "      ") 1 6)
          (gbm:fmt (nth 1 rec) 2) "  "
          (gbm:fmt (nth 2 rec) 2) "  "
          (gbm:fmt (nth 3 rec) 2) "  "
          (gbm:fmt (nth 4 rec) 2) "  "
          (gbm:fmt (nth 5 rec) 2) "  "
          (gbm:fmt (nth 6 rec) 2) "  "
          (gbm:fmt (nth 7 rec) 2) "  "
          (nth 8 rec) " / " (nth 9 rec))))
      (princ "\n─────────────────────────────────────────")
      (princ (strcat "\n 小計 (" (itoa (length *GBM-DATA*)) " 空間):"))
      (princ (strcat "\n   Σ天花板 = " (gbm:fmt a1sum 2) " m²"))
      (princ (strcat "\n   Σ內部裝修 = " (gbm:fmt a2sum 2) " m²"))
      (princ (strcat "\n   Σ天花綠建材 Gi1 = " (gbm:fmt g1sum 2) " m²"))
      (princ (strcat "\n   Σ牆面綠建材 Gi4 = " (gbm:fmt g4sum 2) " m²"))))
  (princ))

;;; ------------------ GBM-EDIT ------------------

(defun c:GBM-EDIT ( / id rec g1_pct g4_pct ai1 ai2)
  (setq id (getstring "\n要修改的空間編號: "))
  (setq rec (gbm:find-record id))
  (cond
    ((null rec) (princ (strcat "\n找不到空間 " id "。")))
    (T
      (setq ai1 (nth 1 rec))
      (setq ai2 (nth 5 rec))
      (princ (strcat "\n空間 " id " 當前: Ai,1=" (gbm:fmt ai1 2)
                     "  Ai,2=" (gbm:fmt ai2 2)
                     "  Gi1=" (gbm:fmt (nth 6 rec) 2)
                     "  Gi4=" (gbm:fmt (nth 7 rec) 2)))
      (setq g1_pct (getreal "\n新天花板綠建材 % <保持原值>: "))
      (setq g4_pct (getreal "\n新牆面綠建材 % <保持原值>: "))
      (gbm:replace-record id
        (list id ai1 (nth 2 rec) (nth 3 rec) (nth 4 rec) ai2
              (if g1_pct (* ai1 (/ g1_pct 100.0)) (nth 6 rec))
              (if g4_pct (* ai2 (/ g4_pct 100.0)) (nth 7 rec))
              (nth 8 rec) (nth 9 rec) (nth 10 rec)))
      (princ (strcat "\n✔ 已更新空間 " id "。"))))
  (princ))

;;; ------------------ GBM-DELETE ------------------

(defun c:GBM-DELETE ( / id)
  (setq id (getstring "\n要刪除的空間編號: "))
  (setq *GBM-DATA*
    (vl-remove-if '(lambda (r) (equal (nth 0 r) id)) *GBM-DATA*))
  (princ (strcat "\n已刪除空間 " id "。剩餘 " (itoa (length *GBM-DATA*)) " 筆。"))
  (princ))

;;; ------------------ GBM-CLEAR ------------------

(defun c:GBM-CLEAR ()
  (setq *GBM-DATA* '())
  (princ "\n已清空所有綠建材資料。")
  (princ))

;;; ------------------ GBM-EXPORT ------------------

(defun c:GBM-EXPORT ( / fn fp rec)
  (cond
    ((null *GBM-DATA*) (princ "\n尚無資料可匯出。"))
    (T
      (setq fn (getfiled "儲存綠建材 CSV" "green_materials.csv" "csv" 1))
      (cond
        ((null fn) (princ "\n已取消。"))
        (T
          (setq fp (open fn "w"))
          ;; Write UTF-8 BOM
          (write-char 239 fp) (write-char 187 fp) (write-char 191 fp)
          (write-line "space_id,Ai1_m2,Li_m,k,H2_m,Ai4_m2,Ai2_m2,Gi1_m2,Gi4_m2,ceiling_material,wall_material" fp)
          (foreach rec *GBM-DATA*
            (write-line
              (strcat
                (nth 0 rec) ","
                (gbm:fmt (nth 1 rec) 4) ","
                (gbm:fmt (nth 2 rec) 4) ","
                (gbm:fmt (nth 3 rec) 4) ","
                (gbm:fmt (nth 4 rec) 4) ","
                (gbm:fmt (nth 10 rec) 4) ","
                (gbm:fmt (nth 5 rec) 4) ","
                (gbm:fmt (nth 6 rec) 4) ","
                (gbm:fmt (nth 7 rec) 4) ","
                (nth 8 rec) ","
                (nth 9 rec))
              fp))
          (close fp)
          (princ (strcat "\n✔ 已匯出 " (itoa (length *GBM-DATA*)) " 筆綠建材資料"))
          (princ (strcat "\n  檔案: " fn))
          (princ "\n  → 請把此 CSV 放到 DWG 同資料夾，GreenBuildingGBF.exe 掃描時會自動帶入。")))))
  (princ))

;;; ------------------ GBM-LOAD ------------------
;;; Load data back from a CSV (so user can resume editing)

(defun gbm:split-csv (line / result cur c i L in-quote)
  (setq result '() cur "" i 1 L (strlen line) in-quote nil)
  (while (<= i L)
    (setq c (substr line i 1))
    (cond
      ((= c "\"") (setq in-quote (not in-quote)))
      ((and (= c ",") (not in-quote))
        (setq result (append result (list cur))) (setq cur ""))
      (T (setq cur (strcat cur c))))
    (setq i (1+ i)))
  (setq result (append result (list cur)))
  result)

(defun c:GBM-LOAD ( / fn fp line fields first-line count)
  (setq fn (getfiled "載入綠建材 CSV" "green_materials.csv" "csv" 0))
  (cond
    ((null fn) (princ "\n已取消。"))
    (T
      (setq fp (open fn "r"))
      (cond
        ((null fp) (princ "\n無法開啟檔案。"))
        (T
          (setq *GBM-DATA* '() count 0 first-line T)
          (while (setq line (read-line fp))
            ;; Strip UTF-8 BOM from first line if present
            (if first-line
              (progn
                (if (and (> (strlen line) 2)
                         (= (ascii (substr line 1 1)) 239))
                  (setq line (substr line 4)))
                (setq first-line nil))
              ;; Skip header and parse data rows
              (progn
                (setq fields (gbm:split-csv line))
                (if (>= (length fields) 11)
                  (progn
                    (setq *GBM-DATA*
                      (append *GBM-DATA*
                        (list (list
                          (nth 0 fields)            ; id
                          (atof (nth 1 fields))     ; Ai1
                          (atof (nth 2 fields))     ; Li
                          (atof (nth 3 fields))     ; k
                          (atof (nth 4 fields))     ; H2
                          (atof (nth 6 fields))     ; Ai2
                          (atof (nth 7 fields))     ; Gi1
                          (atof (nth 8 fields))     ; Gi4
                          (nth 9 fields)            ; ceiling_mat
                          (nth 10 fields)           ; wall_mat
                          (atof (nth 5 fields)))))) ; Ai4
                    (setq count (1+ count)))))))
          (close fp)
          (princ (strcat "\n✔ 已載入 " (itoa count) " 筆綠建材資料。"))))))
  (princ))

;;; ------------------ GBM-TABLE (draw table in drawing) ------------------

(defun c:GBM-TABLE ( / pt x y dy th sw rec)
  (cond
    ((null *GBM-DATA*) (princ "\n尚無資料。"))
    (T
      (setq pt (getpoint "\n指定表格左上角插入點: "))
      (cond
        ((null pt) (princ "\n取消。"))
        (T
          (setq x (car pt) y (cadr pt))
          (setq th 200)     ; text height (drawing units)
          (setq dy (* th 1.8))
          ;; Title
          (command "._TEXT" (list x y) (* th 1.5) 0
            (strcat "綠建材空間檢討表  (" (itoa (length *GBM-DATA*)) " 空間)"))
          (setq y (- y (* dy 1.5)))
          ;; Header row
          (command "._TEXT" (list x y) th 0
            "ID     Ai,1     Li      k    H2    Ai,4     Ai,2     Gi1      Gi4     天花材料      牆面材料")
          (setq y (- y dy))
          ;; Data rows
          (foreach rec *GBM-DATA*
            (command "._TEXT" (list x y) th 0
              (strcat
                (substr (strcat (nth 0 rec) "      ") 1 6)
                (gbm:fmt (nth 1 rec) 2) "  "
                (gbm:fmt (nth 2 rec) 2) "  "
                (gbm:fmt (nth 3 rec) 2) "  "
                (gbm:fmt (nth 4 rec) 2) "  "
                (gbm:fmt (nth 10 rec) 2) "  "
                (gbm:fmt (nth 5 rec) 2) "  "
                (gbm:fmt (nth 6 rec) 2) "  "
                (gbm:fmt (nth 7 rec) 2) "  "
                (substr (strcat (nth 8 rec) "            ") 1 12) "  "
                (nth 9 rec)))
            (setq y (- y dy)))
          (princ (strcat "\n✔ 已繪製表格，共 " (itoa (length *GBM-DATA*)) " 列。"))))))
  (princ))

;;; ------------------ GBM (help) ------------------

(defun c:GBM ()
  (princ "\n========== 綠建材計算工具 ==========")
  (princ "\n  GBM-ADD       新增空間（選封閉線自動抓面積／周長）")
  (princ "\n  GBM-HATCH     統計剖面線（HATCH）面積")
  (princ "\n  GBM-LIST      列出所有空間 + 總計")
  (princ "\n  GBM-EDIT      修改綠建材比例")
  (princ "\n  GBM-DELETE    刪除空間")
  (princ "\n  GBM-CLEAR     清空資料")
  (princ "\n  GBM-EXPORT    匯出 CSV (給 GreenBuildingGBF.exe)")
  (princ "\n  GBM-LOAD      從 CSV 載回資料")
  (princ "\n  GBM-TABLE     在圖面上繪製檢討表")
  (princ "\n======================================")
  (princ (strcat "\n目前已記錄 " (itoa (length *GBM-DATA*)) " 個空間"))
  (princ))

(princ "\n[GBM 綠建材 LISP 已載入]  輸入 GBM 查看指令")
(princ)
