@echo off
title Gooduelle Performance Demo - Port 5173
cd /d "%~dp0"
npx --yes serve -l 5173 .
