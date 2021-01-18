;;------------------------------------------------------------------------
;;                Time-stamp: "2021-01-18 17:10:37 jpdur"
;; ----------------------------------------------------------------------

;; Function which asks for the name of the xls/xlsx file and then insert the table at point
(defun my/xl-to-org-table (filename)
       "Read an XL spreadsheet and insert a table into a Org buffer."
       (interactive "FFind file: ")
       (insert (shell-command-to-string (concat "Powershell XLTable2String('" filename "')") ) )
)

;; Example of key mapping 
(define-key org-mode-map (kbd "C-c x l o") 'my/xl-to-org-table)
