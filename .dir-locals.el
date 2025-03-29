((nil . ((org-export-with-toc . nil)
	 (eval . (let ((proot (project-root (project-current)))) ;vc-root-dir function doesn't work maybe because vc elisp files have not been loaded.
		   (setq org-publish-project-alist
			 `(("ExcelVba-html"
			    ;; Path to your org files.
			    :base-directory ,(concat proot "docs/org/")
			    :base-extension "org"

			    ;; Path to your Jekyll project.
			    :publishing-directory ,(concat proot "docs/")
			    :recursive t
			    :publishing-function org-html-publish-to-html
			    :headline-levels 4
			    :html-extension "html"
			    :section-numbers nil
			    :with-tags nil
			    :body-only t ;; Only export section between <body> </body>
			    )))))
	 )))
