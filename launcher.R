# ============================================================================
#  Microplastic Data Processor - launcher for PCs that block .bat files
#
#  Work PCs with AppLocker script rules block "Start Microplastic Processor.bat"
#  ("This program is blocked by group policy"). A normal Windows shortcut that
#  starts Rscript.exe is not a script and still works. The shortcut downloads
#  and runs this file from GitHub, so this launcher updates automatically too.
#
#  ONE-TIME SETUP (see README, "For users"): open R from the Start menu, paste
#
#    source("https://raw.githubusercontent.com/KristofM854/excel-processor-ftir-raman/main/launcher.R")
#
#  and press Enter. That installs what is needed, copies the shortcut target
#  to the clipboard and explains how to create the desktop shortcut.
#
#  DAILY USE: the desktop shortcut runs this file with Rscript, which starts
#  the latest version of the app from GitHub.
# ============================================================================

local({
  owner <- "KristofM854"
  repo  <- "excel-processor-ftir-raman"
  ref   <- "main"
  launcher_url <- sprintf("https://raw.githubusercontent.com/%s/%s/%s/launcher.R", owner, repo, ref)

  # The desktop shortcut runs Rscript (non-interactive); pasting the setup line
  # into R / RStudio is interactive.
  setup_mode <- interactive()

  banner <- function(...) {
    line <- strrep("=", 72)
    message("\n", line, "\n", paste0("  ", c(...), collapse = "\n"), "\n", line, "\n")
  }

  # Show an error and keep the console window open so it can be read.
  fail <- function(...) {
    banner(...)
    if (setup_mode) stop("Setup did not finish (see the message above).", call. = FALSE)
    message("Press Enter to close this window.")
    invisible(readLines("stdin", n = 1L))
    quit(save = "no", status = 1L)
  }

  # ---- CRAN mirror + writable per-user library ------------------------------
  # R's own library under C:/Program Files is usually read-only on work PCs.
  if (is.null(getOption("repos")) || identical(unname(getOption("repos")["CRAN"]), "@CRAN@")) {
    options(repos = c(CRAN = "https://cloud.r-project.org"))
  }
  user_lib <- Sys.getenv("R_LIBS_USER")
  if (nzchar(user_lib)) {
    user_lib <- path.expand(strsplit(user_lib, .Platform$path.sep, fixed = TRUE)[[1]][1])
    if (!dir.exists(user_lib)) dir.create(user_lib, recursive = TRUE, showWarnings = FALSE)
    if (dir.exists(user_lib)) .libPaths(c(user_lib, .libPaths()))
  }

  # ---- shiny ----------------------------------------------------------------
  if (!requireNamespace("shiny", quietly = TRUE)) {
    message("Installing the R package 'shiny' (first start only, this can take a few minutes)...")
    try(utils::install.packages("shiny"))
    if (!requireNamespace("shiny", quietly = TRUE)) {
      fail("The R package 'shiny' could not be installed.",
           "",
           "Check your internet connection and try again. If it keeps failing,",
           "ask IT to install the R package 'shiny'.")
    }
  }

  # ---- Setup mode: give the user the shortcut target -------------------------
  if (setup_mode) {
    rscript <- file.path(R.home("bin"), "Rscript.exe")
    if (!file.exists(rscript)) rscript <- file.path(dirname(R.home("bin")), "Rscript.exe")
    rscript <- normalizePath(rscript, winslash = "\\", mustWork = FALSE)

    # Kept short: the Windows shortcut dialog accepts at most 259 characters.
    # If the launcher cannot be downloaded, show why and wait for Enter.
    expr   <- sprintf("tryCatch(source('%s'),error=function(e){message('Offline? ',e$message);x=readLines('stdin',1)})",
                      launcher_url)
    target <- sprintf('"%s" -e "%s"', rscript, expr)
    # Long R paths (e.g. a per-user install under AppData): use the short
    # 8.3 form of the path, which Windows shortcuts accept as well.
    if (nchar(target) > 259L && .Platform$OS.type == "windows") {
      short <- tryCatch(utils::shortPathName(rscript), error = function(e) rscript)
      target <- sprintf('"%s" -e "%s"', short, expr)
    }

    copied <- FALSE
    if (.Platform$OS.type == "windows") {
      copied <- isTRUE(tryCatch({ utils::writeClipboard(target); TRUE }, error = function(e) FALSE))
    }

    banner(
      "Setup done. Now create the desktop shortcut (one time only):",
      "",
      "  1. Right-click an empty spot on the desktop -> New -> Shortcut.",
      if (copied) "  2. Paste (Ctrl+V) into the location field - it is already copied."
      else        "  2. Paste the line printed below into the location field.",
      "  3. Click Next, name it 'Microplastic Processor', click Finish.",
      "",
      "Double-click the shortcut to start the app. Every start loads the",
      "latest version automatically. Keep the black window open while you",
      "work; close it to stop the app."
    )
    cat(target, "\n\n")
    if (nchar(target) > 259L) {
      message("Note: this line is ", nchar(target), " characters long; the shortcut dialog may cut it off ",
              "(limit 259). Ask the app maintainer for help.")
    }
    return(invisible(target))
  }

  # ---- Normal start (from the shortcut) --------------------------------------
  message("Starting the Microplastic Data Processor (latest version from GitHub)...")
  message("Keep this window open while you use the app. Close it to stop the app.\n")

  ok <- tryCatch({
    shiny::runGitHub(repo, owner, ref = ref, launch.browser = TRUE)
    TRUE
  }, error = function(e) {
    fail("The app could not be started:",
         paste0("  ", conditionMessage(e)),
         "",
         "If this mentions a download or connection problem, check your network",
         "or VPN connection and try again. Otherwise send a screenshot of this",
         "window to the app maintainer.")
  })
  invisible(ok)
})
