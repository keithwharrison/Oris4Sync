SUBDIRS = ubuntu-mono-dark \
	ubuntu-mono-light

theme = hicolor
system_themedir = $(datadir)/icons/$(theme)
app_themedir = $(pkgdatadir)/icons/$(theme)

system_theme_icons = \
	apps,folder-cmissync-16.png \
	apps,folder-cmissync-22.png	\
	apps,folder-cmissync-24.png \
	apps,folder-cmissync-256.png \
	apps,folder-cmissync-32.png \
	apps,folder-cmissync-48.png

app_theme_icons = \
	animations,process-working-22.png \
	places,folder-cmissync-16.png \
	places,folder-cmissync-22.png \
	places,folder-cmissync-24.png \
	places,folder-cmissync-256.png \
	places,folder-cmissync-32.png \
	places,folder-cmissync-48.png \
	status,process-syncing-i-24.png \
	status,process-syncing-ii-24.png \
	status,process-syncing-iii-24.png \
	status,process-syncing-iiii-24.png \
	status,process-syncing-iiiii-24.png \
	status,process-syncing-error-24.png	\
	status,document-added-12.png \
	status,document-edited-12.png \
	status,document-deleted-12.png \
	status,document-moved-12.png \
	status,list-point-16.png

install_icon_exec = $(top_srcdir)/build/icon-theme-installer \
	-t "$(theme)" \
	-s "$(srcdir)" \
	-d "x$(DESTDIR)" \
	-m "$(mkinstalldirs)" \
	-x "$(INSTALL_DATA)"

install-data-local:
	@-$(install_icon_exec) -i -b "$(system_themedir)" $(system_theme_icons)
	@-$(install_icon_exec) -i -b "$(app_themedir)" $(app_theme_icons)

uninstall-hook:
	@-$(install_icon_exec) -u -b "$(system_themedir)" $(system_theme_icons)
	@-$(install_icon_exec) -u -b "$(app_themedir)" $(app_theme_icons)

MAINTAINERCLEANFILES = Makefile.in
EXTRA_DIST = $(wildcard *.png *.svg)
