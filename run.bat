@echo off
rem =====================================================================
rem SENTINEL STUDIO™ — PORTABLE DESKTOP LAUNCHER
rem Architect: Patrick Jay F. Tanap, CPA, CICA
rem Launches standalone unpacked bundle with Vulkan & WebGPU acceleration
rem =====================================================================
start "" "%~dp0Sentinel Studio.exe" --enable-features=Vulkan,WebGPU --enable-unsafe-webgpu
