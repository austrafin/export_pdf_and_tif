; Matti Syrjanen Distribution Power Design 2017
; Version 1.1
; Changes
; 1. Altered CHANGEFOLDER command so that it creates a text file with the file path to the export folder.
;    This file is loaded (if exists) when and the file path is used when opening the drawing.
; 2. Altered the CHANGEFOLDER command so that the user can navigate to the export folder using a graphical file dialog.
; 3. If the current export path is not valid, the user is prompted to select a new path instead of alerting about invalid path.
; This program automatically exports PDF and TIF files in their correct folders
; ******************************************************************************************************************

; Returns the file name of the file that stores the PDF and TIF export filepath
(defun export_filename () (strcat (getvar 'DWGPREFIX) "export_filepath.txt"))

(defun offset_sapn () 7) ; For trimming the notification type and file extension
(defun offset_other () 4) ; For trimming the file extension

; paper sizes
(defun a0_pdf () "ISO full bleed A0 (841.00 x 1189.00 MM)")
(defun a1_pdf () "ISO full bleed A1 (594.00 x 841.00 MM)")
(defun a2_pdf () "ISO full bleed A2 (420.00 x 594.00 MM)")
(defun a3_pdf () "ISO full bleed A3 (297.00 x 420.00 MM)")
(defun a4_pdf () "ISO full bleed A4 (210.00 x 297.00 MM)")

(defun a0_tif () "A0 841 x 1189 mm")
(defun a1_tif () "A1 594 x 841 mm")
(defun a2_tif () "A2")
(defun a3_tif () "A3")
(defun a4_tif () "A4")

; sheet upper corner coordinates
(defun a0_uc_L () "1189,841")
(defun a1_uc_L () "841,594")
(defun a2_uc_L () "594,420")
(defun a3_uc_L () "420,297")
(defun a4_uc_L () "297,210")

(defun a0_uc_P () "841,1189")
(defun a1_uc_P () "594,841")
(defun a2_uc_P () "420,594")
(defun a3_uc_P () "297,420")
(defun a4_uc_P () "210,297")

; Returns the filepath and the file name with the extension omitted. If SAPN drawing, also crops the notification type.
(defun get_filepath (offset)
	(setq filename (getvar 'DWGNAME))
	(setq fn_len (strlen filename))
	(strcat *folder* (substr filename 1 (- fn_len offset)))
)

(defun extract_notification_type ()
	(setq filename (getvar 'DWGNAME))
	(substr filename (- (strlen filename) 6) 3) ; include dash
)

(defun check_exists ()
	(vl-file-directory-p *folder*)
)

; Credit to Lee Mac
(defun LM:DirectoryDialog ( msg dir flag / Shell Fold Self Path )
	(vl-catch-all-apply
		(function
			(lambda ( / ac HWND )
				(if
					(setq Shell (vla-getInterfaceObject (setq ac (vlax-get-acad-object)) "Shell.Application")
						HWND  (vl-catch-all-apply 'vla-get-HWND (list ac))
						Fold  (vlax-invoke-method Shell 'BrowseForFolder (if (vl-catch-all-error-p HWND) 0 HWND) msg flag dir)
					)
					(setq Self (vlax-get-property Fold 'Self)
					Path (vlax-get-property Self 'Path)
					Path (vl-string-right-trim "\\" (vl-string-translate "/" "\\" Path))
					)
				)
			)
		)
	)
	(if Self  (vlax-release-object  Self))
	(if Fold  (vlax-release-object  Fold))
	(if Shell (vlax-release-object Shell))
	Path
)

(defun print_pdf (revision sheet_number notification_type offset paper_size orientation upper_right)
	(cond
		((vl-file-directory-p *folder*)
	(command "-plot"
		"yes" ; Detailed plot configuration? [Yes/No] <No>:
		"" ; Enter a layout name or [?] <>: 
		"DWG To PDF.pc3" ; Enter an output device name or [?] <None>:
		paper_size ; Enter paper size or [?] <>:
		"Millimeters" ; Enter paper units <Millimeters>
		orientation ; Enter drawing orientation [Portrait Landscape]
		"no" ; Plot upside down? [Yes/No] <>:
		"Window" ; Enter plot area [Display/Extents/Layout/View/Window] <>:
		"0,0" ; Enter lower left corner of window <0.000000,0.000000>:
		upper_right ; Enter upper right corner of window <>:
		"1=1" ; Enter plot scale (Plotted Millimeters=Drawing Units) or [Fit] <1=1>:
		"Center" ; Enter plot offset (x,y) or [Center] <Center>:
		"Yes" ; Plot with plot styles? [Yes/No] <Yes>:
		"acad.ctb" ; Enter plot style table name or [?] (enter . for none) <>:
		"Yes" ; Plot with lineweights? [Yes/No] <>:
		"Yes" ; Scale lineweights with plot scale? [Yes/No] <>:
		"No" ; Plot paper space first? [Yes/No] <>:
		"No" ; Hide paperspace objects? [Yes/No] <No>:
		;filepath\drawing_number(without file extension and notification_type)-sheet_no-notificationtype-file_extension
		(strcat (get_filepath offset) sheet_number notification_type revision ".pdf")
		"no" ; Save changes to page setup [Yes/No]? <N>
		"Yes" ; Proceed with plot [Yes/No] <Y>:
	))
	)
)

(defun print_tif (revision sheet_number notification_type offset paper_size orientation upper_right)

	(command "-plot"
		"yes" ; Detailed plot configuration? [Yes/No] <No>:
		"" ; Enter a layout name or [?] <>: 
		"TIFF Image Printer 11.0" ; Enter an output device name or [?] <None>:
		paper_size ; Enter paper size or [?] <>:
		"Millimeters" ; Enter paper units <Millimeters>
		orientation ; Enter drawing orientation [Portrait Landscape]
		"no" ; Plot upside down? [Yes/No] <>:
		"Window" ; Enter plot area [Display/Extents/Layout/View/Window] <>:
		"0,0" ; Enter lower left corner of window <0.000000,0.000000>:
		upper_right ; Enter upper right corner of window <>:
		"1=1" ; Enter plot scale (Plotted Millimeters=Drawing Units) or [Fit] <1=1>:
		"Center" ; Enter plot offset (x,y) or [Center] <Center>:
		"Yes" ; Plot with plot styles? [Yes/No] <Yes>:
		"acad.ctb" ; Enter plot style table name or [?] (enter . for none) <>:
		"Yes" ; Plot with lineweights? [Yes/No] <>:
		"Yes" ; Scale lineweights with plot scale? [Yes/No] <>:
		"No" ; Plot paper space first? [Yes/No] <>:
		"No" ; Hide paperspace objects? [Yes/No] <No>:
		"Yes" ; Write the plot to a file [Yes/No] <>: (this is the only additional input to print pdf)
		;filepath\drawing_number(without file extension and notification_type)-sheet_no-notificationtype-file_extension
		(strcat (get_filepath offset) sheet_number notification_type revision)
		"no" ; Save changes to page setup [Yes/No]? <N>
		"Yes" ; Proceed with plot [Yes/No] <Y>:
	)
	
)

(defun prompt_for_revision ()
	(setq revision (getstring T "Revision: "))
	(cond ((/= revision "") (setq revision (strcat "-" revision)))
		  (revision)
	)
)

(defun prompt_for_sheet_number ()
	(setq sheet_no (getstring T "Sheet number: "))
	(cond ((/= sheet_no "")
		(cond
			((< (strlen sheet_no) 2) (setq sheet_no (strcat "-0" sheet_no)))
			((setq sheet_no (strcat "-" sheet_no))))
		)
		(sheet_no)
	)
)

(defun check_exists ()
	(cond
		((= (vl-file-directory-p *folder*) nil)
				(c:CHANGEFOLDER))
		(T)
	)
)

(defun pdf_and_tif (notification_type offset paper_size_pdf paper_size_tif orientation upper_right)
	(cond
		((check_exists)
		(setq revision (prompt_for_revision))
		(setq sheet_number (prompt_for_sheet_number))
		(print_pdf revision sheet_number notification_type offset paper_size_pdf orientation upper_right)
		(print_tif revision sheet_number notification_type offset paper_size_tif orientation upper_right))
	)
)

(defun pdf_only (notification_type offset paper_size orientation upper_corner)
	(cond
		((check_exists)
		(print_pdf (prompt_for_revision) (prompt_for_sheet_number) notification_type offset paper_size orientation upper_corner))
	)
)

(defun tif_only (notification_type offset paper_size orientation upper_corner)
	(cond
		((check_exists)
		(print_tif (prompt_for_revision) (prompt_for_sheet_number) notification_type offset paper_size orientation upper_corner))
	)
)


; *************************commands**********************************

; For changing the filepath to the folder where files are saved
(defun c:CHANGEFOLDER ()
	(setq *folder* (strcat (LM:DirectoryDialog "Select the export file path" "" 1) "\\"))
	(setq f (open (export_filename) "w"))
	(write-line *folder* f)
	(close f)
)

; For SA Power Networks drawings

; These command can be used for printing PDFs in landscape orientation
(defun c:EPDFA0LSAPN ()
	(pdf_only (extract_notification_type) (offset_sapn) (a0_pdf) "Landscape" (a0_uc_L))
)

(defun c:EPDFA1LSAPN ()
	(pdf_only (extract_notification_type) (offset_sapn) (a1_pdf) "Landscape" (a1_uc_L))
)

(defun c:EPDFA2LSAPN ()
	(pdf_only (extract_notification_type) (offset_sapn) (a2_pdf) "Landscape" (a2_uc_L))
)

(defun c:EPDFA3LSAPN ()
	(pdf_only (extract_notification_type) (offset_sapn) (a3_pdf) "Landscape" (a3_uc_L))
)

(defun c:EPDFA4LSAPN ()
	(pdf_only (extract_notification_type) (offset_sapn) (a4_pdf) "Landscape" (a4_uc_L))
)

; These command can be used for printing TIFs in landscape orientation
(defun c:ETIFA0LSAPN ()
	(tif_only (extract_notification_type) (offset_sapn) (a0_tif) "Landscape" (a0_uc_L))
)

(defun c:ETIFA1LSAPN ()
	(tif_only (extract_notification_type) (offset_sapn) (a1_tif) "Landscape" (a1_uc_L))
)

(defun c:ETIFA2LSAPN ()
	(tif_only (extract_notification_type) (offset_sapn) (a2_tif) "Landscape" (a2_uc_L))
)

(defun c:ETIFA3LSAPN ()
	(tif_only (extract_notification_type) (offset_sapn) (a3_tif) "Landscape" (a3_uc_L))
)

(defun c:ETIFA4LSAPN ()
	(tif_only (extract_notification_type) (offset_sapn) (a4_tif) "Landscape" (a4_uc_L))
)

; These command can be used for printing both, PDFs and TIFs in landscape orientation
(defun c:EPDFTIFA0LSAPN ()
	(pdf_and_tif (extract_notification_type) (offset_sapn) (a0_pdf) (a0_tif) "Landscape" (a0_uc_L))
)

(defun c:EPDFTIFA1LSAPN ()

	(pdf_and_tif (extract_notification_type) (offset_sapn) (a1_pdf) (a1_tif) "Landscape" (a1_uc_L))
	
)

(defun c:EPDFTIFA2LSAPN ()
	(pdf_and_tif (extract_notification_type) (offset_sapn) (a2_pdf) (a2_tif) "Landscape" (a2_uc_L))
)

(defun c:EPDFTIFA3LSAPN ()
	(pdf_and_tif (extract_notification_type) (offset_sapn) (a3_pdf) (a3_tif) "Landscape" (a3_uc_L))
)

(defun c:EPDFTIFA4LSAPN ()
	(pdf_and_tif (extract_notification_type) (offset_sapn) (a4_pdf) (a4_tif) "Landscape" (a4_uc_L))
)

; These command can be used for printing PDFs in portrait orientation
(defun c:EPDFA0PSAPN ()
	(pdf_only (extract_notification_type) (offset_sapn) (a0_pdf) "Portrait" (a0_uc_P))
)

(defun c:EPDFA1PSAPN ()
	(pdf_only (extract_notification_type) (offset_sapn) (a1_pdf) "Portrait" (a1_uc_P))
)

(defun c:EPDFA2PSAPN ()
	(pdf_only (extract_notification_type) (offset_sapn) (a2_pdf) "Portrait" (a2_uc_P))
)

(defun c:EPDFA3PSAPN ()
	(pdf_only (extract_notification_type) (offset_sapn) (a3_pdf) "Portrait" (a3_uc_P))
)

(defun c:EPDFA4PSAPN ()
	(pdf_only (extract_notification_type) (offset_sapn) (a4_pdf) "Portrait" (a4_uc_P))
)

; These command can be used for printing TIFs in portrait orientation
(defun c:ETIFA0PSAPN ()
	(tif_only (extract_notification_type) (offset_sapn) (a0_tif) "Portrait" (a0_uc_P))
)

(defun c:ETIFA1PSAPN ()
	(tif_only (extract_notification_type) (offset_sapn) (a1_tif) "Portrait" (a1_uc_P))
)

(defun c:ETIFA2PSAPN ()
	(tif_only (extract_notification_type) (offset_sapn) (a2_tif) "Portrait" (a2_uc_P))
)

(defun c:ETIFA3PSAPN ()
	(tif_only (extract_notification_type) (offset_sapn) (a3_tif) "Portrait" (a3_uc_P))
)

(defun c:ETIFA4PSAPN ()
	(tif_only (extract_notification_type) (offset_sapn) (a4_tif) "Portrait" (a4_uc_P))
)

; These command can be used for printing both, PDFs and TIFs in portrait orientation
(defun c:EPDFTIFA0PSAPN ()
	(pdf_and_tif (extract_notification_type) (offset_sapn) (a0_pdf) (a0_tif) "Portrait" (a0_uc_P))
)

(defun c:EPDFTIFA1PSAPN ()
	(pdf_and_tif (extract_notification_type) (offset_sapn) (a1_pdf) (a1_tif) "Portrait" (a1_uc_P))
)

(defun c:EPDFTIFA2PSAPN ()
	(pdf_and_tif (extract_notification_type) (offset_sapn) (a2_pdf) (a2_tif) "Portrait" (a2_uc_P))
)

(defun c:EPDFTIFA3PSAPN ()
	(pdf_and_tif (extract_notification_type) (offset_sapn) (a3_pdf) (a3_tif) "Portrait" (a3_uc_P))
)

(defun c:EPDFTIFA4PSAPN ()
	(pdf_and_tif (extract_notification_type) (offset_sapn) (a4_pdf) (a4_tif) "Portrait" (a4_uc_P))
)

; For other drawings

; These command can be used for printing PDFs in landscape orientation
(defun c:EPDFA0LDPD ()
	(pdf_only "" (offset_other) (a0_pdf) "Landscape" (a0_uc_L))
)

(defun c:EPDFA1LDPD ()
	(pdf_only "" (offset_other) (a1_pdf) "Landscape" (a1_uc_L))
)

(defun c:EPDFA2LDPD ()
	(pdf_only "" (offset_other) (a2_pdf) "Landscape" (a2_uc_L))
)

(defun c:EPDFA3LDPD ()
	(pdf_only "" (offset_other) (a3_pdf) "Landscape" (a3_uc_L))
)

(defun c:EPDFA4LDPD ()
	(pdf_only "" (offset_other) (a4_pdf) "Landscape" (a4_uc_L))
)

; These command can be used for printing TIFs in landscape orientation
(defun c:ETIFA0LDPD ()
	(tif_only "" (offset_other) (a0_tif) "Landscape" (a0_uc_L))
)

(defun c:ETIFA1LDPD ()
	(tif_only "" (offset_other) (a1_tif) "Landscape" (a1_uc_L))
)

(defun c:ETIFA2LDPD ()
	(tif_only "" (offset_other) (a2_tif) "Landscape" (a2_uc_L))
)

(defun c:ETIFA3LDPD ()
	(tif_only "" (offset_other) (a3_tif) "Landscape" (a3_uc_L))
)

(defun c:ETIFA4LDPD ()
	(tif_only "" (offset_other) (a4_tif) "Landscape" (a4_uc_L))
)

; These command can be used for printing both, PDFs and TIFs in landscape orientation
(defun c:EPDFTIFA0LDPD ()
	(pdf_and_tif "" (offset_other) (a0_pdf) (a0_tif) "Landscape" (a0_uc_L))
)

(defun c:EPDFTIFA1LDPD ()
	(pdf_and_tif "" (offset_other) (a1_pdf) (a1_tif) "Landscape" (a1_uc_L))
)

(defun c:EPDFTIFA2LDPD ()
	(pdf_and_tif "" (offset_other) (a2_pdf) (a2_tif) "Landscape" (a2_uc_L))
)

(defun c:EPDFTIFA3LDPD ()
	(pdf_and_tif "" (offset_other) (a3_pdf) (a3_tif) "Landscape" (a3_uc_L))
)

(defun c:EPDFTIFA4LDPD ()
	(pdf_and_tif "" (offset_other) (a4_pdf) (a4_tif) "Landscape" (a4_uc_L))
)

; These command can be used for printing PDFs in portrait orientation
(defun c:EPDFA0PDPD ()
	(pdf_only "" (offset_other) (a0_pdf) "Portrait" (a0_uc_P))
)

(defun c:EPDFA1PDPD ()
	(pdf_only "" (offset_other) (a1_pdf) "Portrait" (a1_uc_P))
)

(defun c:EPDFA2PDPD ()
	(pdf_only "" (offset_other) (a1_pdf) "Portrait" (a1_uc_P))
)

(defun c:EPDFA3PDPD ()
	(pdf_only "" (offset_other) (a3_pdf) "Portrait" (a3_uc_P))
)

(defun c:EPDFA4PDPD ()
	(pdf_only "" (offset_other) (a4_pdf) "Portrait" (a4_uc_P))
)

; These command can be used for printing TIFs in portrait orientation
(defun c:ETIFA0PDPD ()
	(tif_only "" (offset_other) (a0_tif) "Portrait" (a0_uc_P))
)

(defun c:ETIFA1PDPD ()
	(tif_only "" (offset_other) (a1_tif) "Portrait" (a1_uc_P))
)

(defun c:ETIFA2PDPD ()
	(tif_only "" (offset_other) (a2_tif) "Portrait" (a2_uc_P))
)

(defun c:ETIFA3PDPD ()
	(tif_only "" (offset_other) (a3_tif) "Portrait" (a3_uc_P))
)

(defun c:ETIFA4PDPD ()
	(tif_only "" (offset_other) (a4_tif) "Portrait" (a4_uc_P))
)

; These command can be used for printing both, PDFs and TIFs in portrait orientation
(defun c:EPDFTIFA0PDPD ()
	(pdf_and_tif "" (offset_other) (a0_pdf) (a0_tif) "Portrait" (a0_uc_P))
)

(defun c:EPDFTIFA1PDPD ()
	(pdf_and_tif "" (offset_other) (a1_pdf) (a1_tif) "Portrait" (a1_uc_P))
)

(defun c:EPDFTIFA2PDPD ()
	(pdf_and_tif "" (offset_other) (a2_pdf) (a2_tif) "Portrait" (a2_uc_P))
)

(defun c:EPDFTIFA3PDPD ()
	(pdf_and_tif "" (offset_other) (a3_pdf) (a3_tif) "Portrait" (a3_uc_P))
)

(defun c:EPDFTIFA4PDPD ()
	(pdf_and_tif "" (offset_other) (a4_pdf) (a4_tif) "Portrait" (a4_uc_P))
)


; If the file path to the export folder has not been set, the PDFs and TIFs are exported into Drawings - PDF version folder
; assuming that the drawing file is stored directly in Autocad folder.
(cond ((findfile (export_filename))
	(setq f (open (export_filename) "r"))
	(setq *folder* (read-line f))
	(close f))
	((setq *folder* (strcat (getvar 'DWGPREFIX) "\..\\Drawings - PDF version\\")))
)