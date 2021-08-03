.PHONY: build

# https://nuitka.net/doc/user-manual.html#tutorial-setup-and-build-on-windows
build:
	nuitka \
	  --assume-yes-for-downloads \
	  --onefile \
	  --follow-imports \
	  oaloader.py
