@echo off
REM Sync_WrestlingRobe_to_Git.bat - Mirror WrestlingRobe to GitWrestlingRobe
REM Preserves .git folder in destination

echo ============================================================
echo SYNCING WrestlingRobe to GitWrestlingRobe
echo ============================================================
echo.
echo Source: C:\Users\azt12\OneDrive\Documents\Business\Textile\WrestlingRobe
echo Dest:   C:\Users\azt12\OneDrive\Documents\Business\Textile\GitWrestlingRobe
echo.
echo This will OVERWRITE all files in GitWrestlingRobe (except .git folder)
echo.
pause

echo.
echo Running robocopy...
echo.

robocopy "C:\Users\azt12\OneDrive\Documents\Business\Textile\WrestlingRobe" "C:\Users\azt12\OneDrive\Documents\Business\Textile\GitWrestlingRobe" /E /XD ".git" /NFL /NDL /NP

echo.
echo ============================================================
echo SYNC COMPLETE
echo ============================================================
echo.
echo Files have been copied to GitWrestlingRobe
echo .git folder was preserved (not overwritten)
echo.

echo ============================================================
echo SYNCING DAILY_DOCS AND HIGH-LEVEL-DESIGN TO GIT MAIN BRANCH
echo ============================================================
echo.

REM Copy Daily_Docs to Git repo root
echo Copying Daily_Docs to Git repo root...
robocopy "C:\Users\azt12\OneDrive\Documents\Business\Textile\GitWrestlingRobe\Programming\FabricETL\DEV\Daily_Docs" "C:\Users\azt12\OneDrive\Documents\Business\Textile\GitWrestlingRobe\Daily_Docs" /E /NFL /NDL /NP

REM Copy High-Level-Design to Git repo root
echo Copying High-Level-Design to Git repo root...
robocopy "C:\Users\azt12\OneDrive\Documents\Business\Textile\GitWrestlingRobe\Programming\FabricETL\DEV\High-Level-Design" "C:\Users\azt12\OneDrive\Documents\Business\Textile\GitWrestlingRobe\High-Level-Design" /E /NFL /NDL /NP

echo.
echo Committing to Git Main branch...
cd /d "C:\Users\azt12\OneDrive\Documents\Business\Textile\GitWrestlingRobe"
git checkout Main
git add Daily_Docs/ High-Level-Design/
git commit -m "Update Daily_Docs and High-Level-Design"
git push origin Main

echo.
echo ============================================================
echo MAIN BRANCH COMMIT COMPLETE
echo ============================================================
echo.

echo ============================================================
echo SYNCING TEXTILEVISION TO GIT DOCUMENT-IMAGE-ANALYSIS BRANCH
echo ============================================================
echo.

REM Switch to Document-Image-Analysis branch
cd /d "C:\Users\azt12\OneDrive\Documents\Business\Textile\GitWrestlingRobe"
git checkout Document-Image-Analysis

REM Clear existing content (except .git)
echo Clearing existing branch content...
for /d %%D in (*) do (
    if /i not "%%D"==".git" rd /s /q "%%D"
)
for %%F in (*) do (
    if /i not "%%F"==".gitignore" del /q "%%F"
)

REM Copy TextileVision folder structure to repo root (from WrestlingRobe source, not GitWrestlingRobe)
echo Copying TextileVision to Document-Image-Analysis branch...
robocopy "C:\Users\azt12\OneDrive\Documents\Business\Textile\WrestlingRobe\Programming\FabricETL\DEV\Code_Root\TextileVision" "C:\Users\azt12\OneDrive\Documents\Business\Textile\GitWrestlingRobe" /E /XD "venv_paddle" "ToolDependencies" "ChartVLM" "__pycache__" /XF "nul" "*.pyc" /NFL /NDL /NP

REM Delete temp files
del /q tmpclaude-* 2>NUL

echo.
echo Committing to Document-Image-Analysis branch...
git add -A
git commit -m "Update TextileVision folder"
git push origin Document-Image-Analysis

echo.
echo ============================================================
echo DOCUMENT-IMAGE-ANALYSIS BRANCH COMMIT COMPLETE
echo ============================================================
echo.

echo ============================================================
echo SYNCING SCHOLARSWEEP TO GIT RESEARCH-CORPUS-GENERATION BRANCH
echo ============================================================
echo.

REM Switch to Research-Corpus-Generation branch
cd /d "C:\Users\azt12\OneDrive\Documents\Business\Textile\GitWrestlingRobe"
git checkout Research-Corpus-Generation

REM Clear existing content (except .git and .gitignore)
echo Clearing existing branch content...
for /d %%D in (*) do (
    if /i not "%%D"==".git" rd /s /q "%%D"
)
for %%F in (*) do (
    if /i not "%%F"==".gitignore" del /q "%%F"
)

REM Copy ScholarSweep folder CONTENTS to repo root (from WrestlingRobe source, not GitWrestlingRobe)
echo Copying ScholarSweep contents to Research-Corpus-Generation branch root...
robocopy "C:\Users\azt12\OneDrive\Documents\Business\Textile\WrestlingRobe\Programming\FabricETL\DEV\Code_Root\ScholarSweep" "C:\Users\azt12\OneDrive\Documents\Business\Textile\GitWrestlingRobe" /E /XD "__pycache__" /XF "nul" "*.pyc" /NFL /NDL /NP

REM Create .gitkeep files for empty folders
echo Creating .gitkeep files for Output and Test_Scripts...
echo. > "C:\Users\azt12\OneDrive\Documents\Business\Textile\GitWrestlingRobe\Output\.gitkeep"
if not exist "C:\Users\azt12\OneDrive\Documents\Business\Textile\GitWrestlingRobe\Test_Scripts\BLANK_FOR_GIT" mkdir "C:\Users\azt12\OneDrive\Documents\Business\Textile\GitWrestlingRobe\Test_Scripts\BLANK_FOR_GIT"
echo. > "C:\Users\azt12\OneDrive\Documents\Business\Textile\GitWrestlingRobe\Test_Scripts\BLANK_FOR_GIT\.gitkeep"

REM Delete temp files
del /q tmpclaude-* 2>NUL

echo.
echo Committing to Research-Corpus-Generation branch...
git add -A
git commit -m "Update ScholarSweep folder"
git push origin Research-Corpus-Generation

echo.
echo ============================================================
echo ALL GIT COMMITS COMPLETE
echo ============================================================
echo.
pause
