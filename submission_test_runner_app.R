#############################################################################
# Submission File Local Test Runner (Shiny app)
#
# Purpose:
#   Simulate the FDA reviewer's local environment (where all submission
#   files live under a hardcoded "C:/submission_files/...") by:
#     1) Copying the "final" submission_files folder from the server path
#        to a local/mapped drive on the Citrix machine (drive letter is a
#        parameter, e.g. "Y:", because different users may be assigned
#        different drive letters).
#     2) Scanning all program files for the hardcoded source drive letter
#        (e.g. "C:/...") and replacing it with the target drive letter
#        (e.g. "Y:/...") -- IN THE COPY ONLY. The original "C:/" master
#        version on the server is never touched.
#     3) Running the (now drive-letter-adjusted) programs with Rscript,
#        one at a time, capturing logs and exit codes.
#     4) Checking whether the expected output files were actually produced
#        (exists, non-empty, freshly modified after the run started).
#
# NOTES / OPEN QUESTIONS (intentionally left as configurable parameters so
# we can iterate during testing, per your message):
#   - Whether Rscript needs to be installed locally on the Citrix machine,
#     or whether it's acceptable to point `rscript_path` at a network/
#     server R installation, is left as a parameter (`rscript_path`).
#     Recommendation: use a LOCAL R install on the Citrix machine so the
#     test genuinely mimics an FDA reviewer's local-only environment.
#   - "Y:" in your screenshot is a mapped network drive
#     (\\brick.d51.lilly.com\<username>), not the physical local disk of
#     the Citrix VM. Functionally this is fine -- Windows/R treat mapped
#     drives and local disks identically for file I/O -- but it will be
#     slower than a true local disk, so large copies may take longer.
#   - Output completeness checking uses a naming heuristic: assume each
#     program "foo.R" should produce an output file "foo.<ext>" for ext in
#     `output_extensions` (default rtf/docx/svg).
#     This is a first pass -- refine once we see real output patterns.
#
# Required packages (install via your internal Artifactory-backed CRAN
# mirror per company policy -- do NOT use public CRAN/npm directly):
#   install.packages(c("shiny", "DT", "processx", "shinyjs"))
#############################################################################

library(shiny)
library(bslib)
library(DT)
library(processx)
library(shinyjs)

# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------

# This app detects, at startup, which machine it's actually running on --
# the real Windows/Citrix box, or a Linux Posit Workbench session -- and
# adapts the Copy step and disk-space check accordingly. Everything else
# (path-prefix replace, running the .R programs, output checking) is pure
# file/string logic and behaves identically on both.
IS_WINDOWS <- .Platform$OS.type == "windows"

# Rscript belonging to THIS running R engine (not just whatever "Rscript" is
# first on PATH) -- this way the app always runs programs with the exact R
# installation it's already using, whether that's this Workbench session or
# a local RStudio on the Citrix machine.
DEFAULT_RSCRIPT_PATH <- file.path(R.home("bin"), if (IS_WINDOWS) "Rscript.exe" else "Rscript")
DEFAULT_PROTOCOL_ID <- "j1i_mc_gzbk"

# Auto-detect the CURRENT user's own drive/root, instead of hardcoding one
# person's setup -- different users can have different IT-assigned drive
# letters for their personal network drive (Y:, M:, etc.), and different
# usernames/home directories on Linux. Windows sets HOMEDRIVE automatically
# per logged-in user (parallels Sys.getenv("HOME") on Linux).
DEFAULT_DEST_ROOT <- if (IS_WINDOWS) {
  hd <- Sys.getenv("HOMEDRIVE")
  if (nchar(hd) > 0) hd else "Y:"  # fallback if HOMEDRIVE isn't set for some reason
} else {
  Sys.getenv("HOME")
}

# Standardized folder layout -- not exposed as inputs since every submission
# package follows the same structure: programs/ and output/ live side by
# side directly under the destination root.
PROGRAMS_SUBDIR <- "programs"
OUTPUT_SUBDIR <- "output"
OUTPUT_EXTENSIONS <- c("rtf", "docx", "svg")

# Normalize a path prefix (works for both a bare drive letter like "C:" and
# a full path like "C:/submission_files"): trim whitespace and any trailing
# slash/backslash, so we can reliably build "prefix/" and "prefix\" forms.
normalize_prefix <- function(p) {
  p <- trimws(p)
  sub("[/\\\\]+$", "", p)
}

# Strip a literal prefix from the start of a string, for display purposes.
# Deliberately NOT regex-based (no escaping needed, no risk of the prefix
# containing regex metacharacters that break pattern compilation) -- just a
# plain substring check.
strip_prefix_literal <- function(x, prefix) {
  ifelse(startsWith(x, prefix), substring(x, nchar(prefix) + 1), x)
}

# Order program file paths so ones ending in "_tf" (self-contained, no
# dependency on other programs' outputs) run before the rest.
order_tf_first <- function(paths) {
  base <- tools::file_path_sans_ext(basename(paths))
  is_tf <- grepl("_tf$", base, ignore.case = TRUE)
  c(paths[is_tf], paths[!is_tf])
}

# Find a column name in a data.frame matching `target`, tolerant of case,
# whitespace, and punctuation differences -- LOA files are hand-maintained
# spreadsheets and headers vary slightly from what you'd type from memory.
find_column <- function(df, target) {
  norm <- function(x) tolower(gsub("[^[:alnum:]]", "", trimws(x)))
  nms <- names(df)
  target_n <- norm(target)
  exact <- which(vapply(nms, norm, character(1)) == target_n)
  if (length(exact) > 0) return(nms[exact[1]])
  contains <- which(grepl(target_n, vapply(nms, norm, character(1)), fixed = TRUE))
  if (length(contains) > 0) return(nms[contains[1]])
  NA_character_
}

# Find a sheet name in an Excel file matching `target`, tolerant of case and
# leading/trailing whitespace (mirrors find_column's tolerance for headers).
find_sheet <- function(path, target) {
  sheets <- readxl::excel_sheets(path)
  norm <- function(x) tolower(trimws(x))
  hit <- which(norm(sheets) == norm(target))
  if (length(hit) > 0) return(sheets[hit[1]])
  NA_character_
}

# Read a LOA-style tracking file (.xlsx/.xls/.csv/.tsv) and return the
# program names flagged "Y/Y" in the given status column. For Excel files:
# if `sheet_name` matches an actual sheet, use it; otherwise fall back to
# the first sheet (more flexible across trackers that don't use this exact
# sheet name) -- but always report which sheet was actually used, so a
# silent fallback never looks like it read the sheet you asked for.
read_loa_yy_programs <- function(path, status_col_name, program_col_name, sheet_name = NULL) {
  ext <- tolower(tools::file_ext(path))
  sheet_used <- NA_character_
  df <- if (ext %in% c("xlsx", "xls")) {
    if (!requireNamespace("readxl", quietly = TRUE)) {
      stop("The 'readxl' package is required to read .xlsx/.xls files but is not installed.")
    }
    sheet_to_use <- NULL
    if (!is.null(sheet_name) && nzchar(sheet_name)) {
      sheet_to_use <- find_sheet(path, sheet_name)  # NA if not found -> falls back to first sheet
      if (is.na(sheet_to_use)) sheet_to_use <- NULL
    }
    result <- as.data.frame(readxl::read_excel(path, sheet = sheet_to_use))
    sheet_used <- if (!is.null(sheet_to_use)) sheet_to_use else readxl::excel_sheets(path)[1]
    result
  } else if (ext == "csv") {
    read.csv(path, stringsAsFactors = FALSE, check.names = FALSE)
  } else if (ext %in% c("tsv", "txt")) {
    read.delim(path, stringsAsFactors = FALSE, check.names = FALSE)
  } else {
    stop(sprintf("Unsupported LOA file extension: .%s (expected .xlsx/.xls/.csv/.tsv)", ext))
  }
  
  status_col <- find_column(df, status_col_name)
  program_col <- find_column(df, program_col_name)
  if (is.na(status_col)) {
    stop(sprintf("Could not find a column matching '%s'. Columns found: %s",
                 status_col_name, paste(names(df), collapse = ", ")))
  }
  if (is.na(program_col)) {
    stop(sprintf("Could not find a column matching '%s'. Columns found: %s",
                 program_col_name, paste(names(df), collapse = ", ")))
  }
  
  status_vals <- trimws(as.character(df[[status_col]]))
  yy_rows <- toupper(status_vals) %in% c("Y/Y")
  programs <- trimws(as.character(df[[program_col]][yy_rows]))
  programs <- programs[nzchar(programs) & !is.na(programs)]
  # A cell can list more than one program name (space/comma/semicolon/pipe
  # separated) -- same convention used elsewhere against this tracker file.
  programs <- unlist(strsplit(programs, "[[:space:],;|]+"))
  programs <- trimws(programs)
  programs <- programs[nzchar(programs)]
  list(programs = unique(programs), status_col = status_col, program_col = program_col,
       n_total_rows = nrow(df), sheet_used = sheet_used)
}

# Match LOA-listed program names (often without extension) against the
# actual discovered .R file paths, case-insensitive, ignoring extension.
match_loa_to_files <- function(loa_names, file_paths) {
  file_base <- tolower(tools::file_path_sans_ext(basename(file_paths)))
  loa_base <- tolower(tools::file_path_sans_ext(trimws(loa_names)))
  matched <- file_paths[file_base %in% loa_base]
  unmatched_loa <- loa_names[!(loa_base %in% file_base)]
  list(matched = matched, unmatched_loa_names = unmatched_loa)
}

# Auto-detect the study/protocol folder name from the real "Copy FROM" path
# -- the segment right after the compound code (matches "ly" + digits, e.g.
# "ly3437943", generically -- not hardcoded to one specific compound so this
# works for other studies/compounds too) and before the next "/". Falls back
# to DEFAULT_PROTOCOL_ID if that pattern isn't found in the path.
extract_protocol_id_from_path <- function(path) {
  m <- regmatches(path, regexpr("(?i)ly[0-9]+/[^/]+", path, perl = TRUE))
  if (length(m) == 0 || !nzchar(m)) return(NA_character_)
  sub("(?i)ly[0-9]+/", "", m, perl = TRUE)
}

# Split a hardcoded path into its DRIVE part (what actually varies between
# the FDA machine and our test machine, e.g. "C:") and its RELATIVE
# structure (the study's own folder layout, e.g.
# "submission_files/j1i_mc_gzbk/final") -- the relative part gets mirrored
# literally in the test destination, so it doesn't matter how any given
# program builds its paths (one long literal string, or dic+transfer style,
# or anything else): only the drive letter itself ever needs replacing.
extract_drive_prefix <- function(p) {
  p <- normalize_prefix(p)
  m <- regmatches(p, regexpr("^[A-Za-z]:", p))
  if (length(m) == 0 || nchar(m) == 0) return(p)  # no drive-letter pattern found; treat whole thing as the prefix
  m
}

extract_relative_structure <- function(p) {
  p <- normalize_prefix(p)
  sub("^[A-Za-z]:[/\\\\]*", "", p)
}

# Build the literal forms of a path prefix we might see in source code:
#   fwd        - "C:/submission_files/j1i_mc_gzbk/"  (continues into a longer path)
#   bwd        - "C:\\submission_files\\j1i_mc_gzbk\\" (one literal backslash)
#   end_dquote - "C:/submission_files/j1i_mc_gzbk\""  (the value IS exactly the
#                prefix, e.g. dic <- "C:/submission_files/j1i_mc_gzbk" with
#                nothing after it -- easy to miss if you only check for a
#                trailing slash, but this pattern is just as common when code
#                builds the rest of the path with file.path()/paste0() later)
#   end_squote - same idea, single-quoted string
path_patterns <- function(prefix) {
  prefix <- normalize_prefix(prefix)
  c(fwd = paste0(prefix, "/"), bwd = paste0(prefix, "\\"),
    end_dquote = paste0(prefix, "\""), end_squote = paste0(prefix, "'"))
}

# Recursively list files under `root` whose extension is in `extensions`
# (extensions given without the dot, e.g. c("R","r")).
list_files_by_ext <- function(root, extensions, recursive = TRUE) {
  if (!dir.exists(root)) return(character(0))
  pattern <- paste0("\\.(", paste(extensions, collapse = "|"), ")$")
  list.files(root, pattern = pattern, recursive = recursive, full.names = TRUE, ignore.case = TRUE)
}

# Count occurrences of `pattern` (fixed string) in a single file's content.
count_occurrences <- function(text, pattern) {
  if (nchar(pattern) == 0) return(0)
  # Count non-overlapping fixed-string matches
  m <- gregexpr(pattern, text, fixed = TRUE)[[1]]
  if (identical(m[1], -1L)) return(0)
  length(m)
}

# Read a file fully as one string, trying UTF-8 first and falling back to
# native encoding, since these program files can come from mixed sources.
read_file_text <- function(path) {
  raw <- tryCatch(readChar(path, file.info(path)$size, useBytes = TRUE),
                  error = function(e) NA_character_)
  raw
}

write_file_text <- function(path, text) {
  con <- file(path, open = "wb")
  on.exit(close(con))
  writeChar(text, con, eos = NULL, useBytes = TRUE)
}

# Scan candidate files for how many times each path pattern appears.
scan_for_drive_refs <- function(files, patterns, ignore_case = FALSE) {
  out <- lapply(files, function(f) {
    txt <- read_file_text(f)
    if (is.na(txt)) return(data.frame(file = f, fwd_count = NA, bwd_count = NA, end_count = NA, error = "read_failed", stringsAsFactors = FALSE))
    search_txt <- if (ignore_case) toupper(txt) else txt
    up_if_needed <- function(x) if (ignore_case) toupper(x) else x
    fwd_n <- count_occurrences(search_txt, up_if_needed(patterns[["fwd"]]))
    bwd_n <- count_occurrences(search_txt, up_if_needed(patterns[["bwd"]]))
    end_n <- count_occurrences(search_txt, up_if_needed(patterns[["end_dquote"]])) +
      count_occurrences(search_txt, up_if_needed(patterns[["end_squote"]]))
    data.frame(
      file = f,
      fwd_count = fwd_n,
      bwd_count = bwd_n,
      end_count = end_n,
      error = NA_character_,
      stringsAsFactors = FALSE
    )
  })
  res <- do.call(rbind, out)
  res$total <- ifelse(is.na(res$fwd_count), NA, res$fwd_count + res$bwd_count + res$end_count)
  res
}

# Apply the replacement in place on each file that has matches.
# `patterns` = output of path_patterns(source_prefix); `replacement_prefix` =
# the new path prefix to substitute in (e.g. the local destination root).
apply_drive_replace <- function(files, patterns, replacement_prefix, ignore_case = FALSE) {
  rep_prefix <- normalize_prefix(replacement_prefix)
  rep_fwd <- paste0(rep_prefix, "/")
  rep_bwd <- paste0(rep_prefix, "\\")
  rep_end_dquote <- paste0(rep_prefix, "\"")
  rep_end_squote <- paste0(rep_prefix, "'")
  
  out <- lapply(files, function(f) {
    txt <- read_file_text(f)
    if (is.na(txt)) {
      return(data.frame(file = f, n_replaced = NA, status = "READ_FAILED", stringsAsFactors = FALSE))
    }
    orig <- txt
    search_txt <- if (ignore_case) toupper(txt) else txt
    up_if_needed <- function(x) if (ignore_case) toupper(x) else x
    n_matches <- count_occurrences(search_txt, up_if_needed(patterns[["fwd"]])) +
      count_occurrences(search_txt, up_if_needed(patterns[["bwd"]])) +
      count_occurrences(search_txt, up_if_needed(patterns[["end_dquote"]])) +
      count_occurrences(search_txt, up_if_needed(patterns[["end_squote"]]))
    if (ignore_case) {
      # Case-insensitive replace: use regex with fixed patterns escaped.
      esc <- function(p) gsub("([\\\\.^$|()\\[\\]{}*+?])", "\\\\\\1", p, perl = TRUE)
      txt <- gsub(esc(patterns[["fwd"]]), rep_fwd, txt, ignore.case = TRUE, perl = TRUE)
      txt <- gsub(esc(patterns[["bwd"]]), rep_bwd, txt, ignore.case = TRUE, perl = TRUE)
      txt <- gsub(esc(patterns[["end_dquote"]]), rep_end_dquote, txt, ignore.case = TRUE, perl = TRUE)
      txt <- gsub(esc(patterns[["end_squote"]]), rep_end_squote, txt, ignore.case = TRUE, perl = TRUE)
    } else {
      txt <- gsub(patterns[["fwd"]], rep_fwd, txt, fixed = TRUE)
      txt <- gsub(patterns[["bwd"]], rep_bwd, txt, fixed = TRUE)
      txt <- gsub(patterns[["end_dquote"]], rep_end_dquote, txt, fixed = TRUE)
      txt <- gsub(patterns[["end_squote"]], rep_end_squote, txt, fixed = TRUE)
    }
    if (identical(orig, txt)) {
      return(data.frame(file = f, n_replaced = 0, status = "NO_CHANGE", stringsAsFactors = FALSE))
    }
    ok <- tryCatch({ write_file_text(f, txt); TRUE }, error = function(e) FALSE)
    data.frame(file = f, n_replaced = if (ok) n_matches else NA,
               status = if (ok) "REPLACED" else "WRITE_FAILED", stringsAsFactors = FALSE)
  })
  do.call(rbind, out)
}

# Guess the expected output file(s) for a program using a naming heuristic.
guess_expected_outputs <- function(program_file, output_dir, output_extensions) {
  base <- tools::file_path_sans_ext(basename(program_file))
  candidates <- file.path(output_dir, paste0(base, ".", output_extensions))
  candidates
}

log_line <- function(msg) {
  paste0("[", format(Sys.time(), "%Y-%m-%d %H:%M:%S"), "] ", msg)
}

# Start the folder copy as an async process, OS-appropriate:
#   Windows -> robocopy (mirrors subfolders, has a real /LOG file)
#   Linux   -> `cp -rv`  (robocopy doesn't exist on Linux/Workbench)
# Returns a processx::process object in both cases so the calling code
# (the copy_timer poller) doesn't need to know which OS it's on.
start_copy_process <- function(source, dest, log_path) {
  dir.create(dest, recursive = TRUE, showWarnings = FALSE)
  
  if (IS_WINDOWS) {
    src_win <- normalizePath(source, winslash = "\\", mustWork = FALSE)
    dst_win <- normalizePath(dest, winslash = "\\", mustWork = FALSE)
    return(processx::process$new(
      "robocopy",
      args = c(src_win, dst_win, "/E", "/R:2", "/W:5", paste0("/LOG:", log_path)),
      stdout = "|", stderr = "|"
    ))
  }
  
  # Linux: copy the CONTENTS of source into dest (trailing "/." mirrors
  # robocopy's /E behavior of not nesting an extra folder level).
  src_contents <- paste0(normalize_prefix(source), "/.")
  processx::process$new(
    "cp",
    args = c("-rv", src_contents, dest),
    stdout = log_path, stderr = log_path
  )
}

# robocopy's exit codes 0-7 all mean "success" of some kind; 8+ is an error.
# cp only ever returns 0 (success) or non-zero (error).
copy_exit_is_success <- function(exit_code) {
  if (is.na(exit_code)) return(FALSE)
  if (IS_WINDOWS) exit_code < 8 else exit_code == 0
}

# Build the actual Rscript invocation for one program. When `use_clean_lib`
# is TRUE, don't run the program file directly -- instead write a tiny
# wrapper script that forces .libPaths() to [clean_lib_dir, base R only]
# BEFORE sourcing the real program. include.site = FALSE is required: R's
# .libPaths() silently re-adds the site library by default even when you
# pass an explicit new vector, unless you turn that off.
build_launch_args <- function(program_path, use_clean_lib, clean_lib_dir) {
  if (!isTRUE(use_clean_lib)) {
    return(list(flags = character(0), script = program_path))
  }
  wrapper_path <- tempfile(fileext = ".R")
  writeLines(c(
    sprintf(".libPaths(c(%s, .Library), include.site = FALSE)", deparse(clean_lib_dir)),
    sprintf("source(%s, chdir = TRUE)", deparse(program_path))
  ), wrapper_path)
  list(flags = "--vanilla", script = wrapper_path)
}

# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------

app_theme <- bslib::bs_theme(
  version = 5,
  primary = "#1d4ed8",
  secondary = "#64748b",
  success = "#16a34a",
  warning = "#d97706",
  danger = "#dc2626",
  bg = "#f8fafc",
  fg = "#0f172a",
  "border-radius" = "0.5rem",
  "card-border-color" = "#e2e8f0",
  "navbar-bg" = "#1d4ed8"
)

ui <- bslib::page_sidebar(
  title = "Submission File Local Test Runner",
  theme = app_theme,
  useShinyjs(),
  
  # Global fix: bslib::card() marks its body as a flex column container
  # (classes html-fill-item/html-fill-container), whose default align-items
  # stretches children -- including plain <button> elements -- to full
  # width. align-self on the item itself is what actually overrides that
  # (width:auto alone does NOT, since 'auto' is exactly what gets stretched).
  tags$style(HTML(".btn { width: auto !important; align-self: flex-start !important; }")),
  tags$style(HTML(paste0(
    "/* Fixed-height cards need their own body to scroll internally, or ",
    "content (tables, logs) overflows and visually overlaps whatever comes ",
    "after it instead of pushing the card taller. bslib also tags DT tables ",
    "(and other outputs) inside a fillable card as flex items that get ",
    "squeezed/stretched to share a fixed height budget -- forcing them back ",
    "to their own natural content height (only for DIRECT children of the ",
    "card body, not the card body itself) is what lets the scrollbar below ",
    "actually work instead of overlapping. */",
    ".card { overflow: hidden; }",
    ".card-body { overflow-y: auto !important; }",
    ".card-body.html-fill-container > .html-fill-item { flex: 0 0 auto !important; height: auto !important; }"
  ))),
  
  sidebar = bslib::sidebar(
    width = 400,
    
    div(
      style = paste0("background:#eef2ff; border:1px solid #dbe4ff; ",
                     "border-radius:8px; padding:14px; margin-bottom:14px;"),
      div(
        style = "display:flex; align-items:center; gap:12px; margin-bottom:10px;",
        div(style = paste0("background:#1d4ed8; color:white; font-weight:700; ",
                           "font-size:1.05rem; width:40px; height:40px; ",
                           "border-radius:8px; display:flex; align-items:center; ",
                           "justify-content:center; flex-shrink:0;"),
            "SF"),
        div(
          div(style = "font-weight:700; font-size:0.98rem; line-height:1.25; color:#0f1f3d;",
              "Submission Test Runner"),
          div(style = "font-size:0.74rem; color:#334155;",
              "Local FDA-environment validation")
        )
      ),
      tags$div(style = paste0("font-size:0.66rem; letter-spacing:0.06em; ",
                              "color:#64748b; text-transform:uppercase; ",
                              "margin-bottom:4px;"),
               "About"),
      p(style = "font-size:0.78rem; color:#1e293b; line-height:1.4; margin-bottom:0;",
        "A local testing tool that validates submission executable programs ",
        "before FDA review.")
    ),
    
    tags$div(style = paste0("font-size:0.68rem; letter-spacing:0.06em; ",
                            "opacity:0.55; text-transform:uppercase; ",
                            "margin-bottom:6px;"),
             "Setup"),
    
    h6("1. Copy FROM (real, current location)"),
    textInput("copy_from_path", NULL,
              value = if (IS_WINDOWS) {
                "Z:/qa/ly3437943/j1i_mc_gzbk/common/documentation/submission/submission_files/j1i_mc_gzbk/final"
              } else {
                "/lillyce/qa/ly3437943/j1i_mc_gzbk/common/documentation/submission/submission_files/j1i_mc_gzbk/final"
              }),
    helpText("Where the deliverable actually sits now. Read-only -- never modified."),
    
    h6("2. LOA file"),
    textInput("loa_path", NULL,
              value = "/lillyce/qa/ly3437943/j1i_mc_gzbk/final/documentation/loa_trackers/gzbk_final_TFL_link.xlsx"),
    helpText("Used in Tab 2 to restrict the run to only Y/Y-flagged programs."),
    
    h6("3. Study/protocol folder (auto-detected)"),
    verbatimTextOutput("hardcoded_path_breakdown"),
    helpText("Pulled from the segment right after the compound code (ly######) in the path above."),
    
    h6("4. Destination drive/root (auto-detected)"),
    verbatimTextOutput("dest_full_preview"),
    helpText(if (IS_WINDOWS) {
      "Your Windows profile's home drive."
    } else {
      "Your home directory (verified to map to the same storage as the Windows drive)."
    }),
    
    bslib::accordion(
      open = FALSE,
      bslib::accordion_panel(
        "Advanced settings",
        h6("R engine"),
        textInput("rscript_path", "Rscript path", value = DEFAULT_RSCRIPT_PATH),
        helpText("Defaults to the R engine running this app. Override only to test a different R installation."),
        checkboxInput("clean_library", "Run with a CLEAN package library (exclude pre-installed extension packages)",
                      value = TRUE),
        helpText("Forces each program to install its own packages instead of using pre-installed ones. ",
                 "Persists within one Run, resets on the next -- still not a substitute for the Windows/Citrix toolchain."),
        actionButton("reset_clean_lib", "Reset clean library", style = "width: auto;"),
        hr(),
        h6("LOA column names (override if your file's headers differ)"),
        textInput("loa_sheet_name", "Sheet name (Excel files only)", value = "TFL Metadata"),
        textInput("loa_status_col", "Status column name", value = "Program for submission/ Executable?"),
        textInput("loa_program_col", "Program-name column name", value = "Program Name"),
        helpText("Matching tolerates case/spacing differences. Falls back to the first sheet ",
                 "if the name isn't found (check Console Log)."),
        hr(),
        h6("Path replacement"),
        textInput("file_extensions", "File extensions to scan/replace", value = "R"),
        checkboxInput("ignore_case", "Ignore case when matching path", value = FALSE),
        checkboxInput("keep_backup", "Keep an untouched backup before replacing", value = FALSE)
      )
    ),
    
    hr(style = "margin: 16px 0 10px;"),
    tags$div(style = paste0("font-size:0.68rem; letter-spacing:0.06em; ",
                            "color:#64748b; text-transform:uppercase; ",
                            "margin-bottom:6px;"),
             "Help"),
    p(style = "font-size:0.78rem; color:#334155; line-height:1.4;",
      "If a step fails, check the Console Log card first -- every action is ",
      "timestamped there. Ask in your team's validation channel for anything ",
      "this app doesn't explain.")
  ),
  
  if (IS_WINDOWS) {
    div(class = "alert alert-success", role = "alert",
        strong("Running on Windows. "),
        "Full environment test -- Copy, Path Replace, Run, and Output Check all ",
        "reflect real Windows/Citrix behavior (including package installation ",
        "and compilation).")
  } else {
    div(class = "alert alert-warning", role = "alert",
        strong("Running on Linux (Posit Workbench). "),
        "This is a LOGIC-ONLY pre-check. It runs the programs end-to-end and can ",
        "catch code bugs, but it does NOT validate package installation, Rtools ",
        "compilation, or any Windows-specific behavior. A pass here is a useful ",
        "early signal, but is not sufficient evidence -- you still need to run ",
        "the full test on the Windows/Citrix machine afterward. (The destination ",
        "path below will literally create a folder named e.g. 'Y:' here, since ",
        "Linux filenames can contain a colon -- this lets the hardcoded ",
        "Windows-style paths in the programs resolve for this logic check.)")
  },
  
  bslib::layout_columns(
    col_widths = c(4, 4, 4),
    
    bslib::card(
      full_screen = TRUE,
      height = "780px",
      bslib::card_header("1. Copy & Prepare for Testing"),
      div(style = "display: flex; flex-direction: row; flex-wrap: wrap; gap: 8px;",
          actionButton("copy_btn", "Copy server folder to Citrix drive", class = "btn-primary btn-sm", style = "width: auto;"),
          actionButton("prepare_btn", "Prepare files for testing", class = "btn-primary btn-sm", style = "width: auto;")
      ),
      helpText("Progress and results appear in the Console Log tab."),
      hr(),
      h6("Files found:"),
      DTOutput("scan_table")
    ),
    
    bslib::card(
      full_screen = TRUE,
      height = "780px",
      bslib::card_header("2. Run Programs"),
      div(style = "display: flex; flex-direction: row; align-items: center; flex-wrap: wrap; gap: 8px 16px;",
          actionButton("refresh_programs_btn", "Refresh program list based on LOA", class = "btn-primary btn-sm", style = "width: auto;"),
          div(style = "white-space: nowrap;", checkboxInput("recurse_subfolders", "Include subfolders", value = FALSE))
      ),
      br(),
      h6(textOutput("programs_select_heading", inline = TRUE)),
      DTOutput("programs_table"),
      uiOutput("loa_unmatched_box"),
      br(),
      div(style = "display: flex; flex-direction: row; flex-wrap: wrap; gap: 8px;",
          actionButton("run_btn", "Run selected programs", class = "btn-primary btn-sm", style = "width: auto;"),
          actionButton("stop_btn", "Stop", class = "btn-danger btn-sm", style = "width: auto;")
      ),
      br(),
      h6("Run status:"),
      DTOutput("run_status_table")
    ),
    
    bslib::card(
      full_screen = TRUE,
      height = "780px",
      bslib::card_header("3. Output & Log"),
      h6("Output Check"),
      div(style = "display: flex; flex-direction: row; flex-wrap: wrap; gap: 8px;",
          actionButton("check_btn", "Check output completeness", class = "btn-primary btn-sm", style = "width: auto;"),
          downloadButton("save_report_btn", "Download report (CSV)", class = "btn-sm", style = "width: auto;")
      ),
      br(),
      DTOutput("completeness_table"),
      
      hr(),
      div(style = "display: flex; flex-direction: row; justify-content: space-between; align-items: center;",
          h6("Console Log", style = "margin-bottom: 0;"),
          actionButton("clear_log_btn", "Clear log", class = "btn-sm", style = "width: auto;")
      ),
      br(),
      tags$style(HTML("#full_log { height: 340px; overflow-y: auto; }")),
      verbatimTextOutput("full_log")
    )
  )
)

# ---------------------------------------------------------------------------
# Server
# ---------------------------------------------------------------------------

server <- function(input, output, session) {
  
  rv <- reactiveValues(
    log = character(0),
    copy_proc = NULL,
    copy_log_path = NULL,
    copy_total_files = NA,
    scan_result = NULL,
    replace_result = NULL,
    programs = character(0),
    run_queue = character(0),
    run_current_proc = NULL,
    run_current_program = NULL,
    run_current_start = NULL,
    run_current_log = NULL,
    run_results = list(),
    run_active = FALSE,
    clean_lib_dir = NULL,
    completeness_result = NULL
  )
  
  add_log <- function(msg) {
    rv$log <- c(rv$log, log_line(msg))
  }
  
  output$full_log <- renderText({
    paste(rv$log, collapse = "\n")
  })
  
  observeEvent(input$clear_log_btn, {
    rv$log <- character(0)
  })
  
  # The ACTUAL physical folder the copied files land in: dest_base_path
  # (just the drive/root, e.g. "Y:") plus the FULL relative structure taken
  # from the hardcoded path (e.g. "submission_files/j1i_mc_gzbk/final"),
  # mirrored literally. This is what makes ANY way a program builds its
  # paths -- one long literal string, dic+transfer style, or anything else
  # -- resolve correctly: only the drive letter itself ever gets replaced,
  # everything after it is preserved exactly as the study already has it.
  # The hardcoded path is fully composed from the study/protocol ID -- not a
  # separate user-facing option. "submission_files" and "final" are the
  # standardized, unchanging convention.
  protocol_id_reactive <- reactive({
    req(input$copy_from_path)
    extracted <- extract_protocol_id_from_path(input$copy_from_path)
    if (is.na(extracted)) DEFAULT_PROTOCOL_ID else extracted
  })
  
  hardcoded_path <- reactive({
    paste0("C:/submission_files/", protocol_id_reactive(), "/final")
  })
  
  effective_dest_root <- reactive({
    file.path(normalize_prefix(DEFAULT_DEST_ROOT), extract_relative_structure(hardcoded_path()))
  })
  
  output$hardcoded_path_breakdown <- renderText({
    sprintf("Full hardcoded path: %s\nDrive to replace: %s  |  Structure to mirror: %s",
            hardcoded_path(),
            extract_drive_prefix(hardcoded_path()),
            extract_relative_structure(hardcoded_path()))
  })
  
  output$dest_full_preview <- renderText({
    paste0("Full test path: ", effective_dest_root())
  })
  
  # -------------------------------------------------------------------------
  # Tab 1: Copy
  # -------------------------------------------------------------------------
  
  observeEvent(input$copy_btn, {
    src <- input$copy_from_path
    dst <- effective_dest_root()
    
    if (!dir.exists(src)) {
      add_log(paste0("ERROR: source path does not exist or is not reachable: ", src))
      showNotification("Source path not found. Check the path and network connectivity.", type = "error")
      return(invisible(NULL))
    }
    
    rv$copy_total_files <- length(list.files(src, recursive = TRUE))
    add_log(sprintf("Starting copy (%s): %s -> %s (%d files found in source; mirroring '%s' under the destination root)",
                    if (IS_WINDOWS) "robocopy" else "cp", src, dst, rv$copy_total_files,
                    extract_relative_structure(hardcoded_path())))
    
    log_path <- file.path(tempdir(), paste0("copy_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".log"))
    rv$copy_log_path <- log_path
    
    proc <- tryCatch({
      start_copy_process(src, dst, log_path)
    }, error = function(e) {
      add_log(paste0("ERROR launching copy: ", conditionMessage(e)))
      NULL
    })
    rv$copy_proc <- proc
  })
  
  # Poll the copy process every second while it's running
  copy_timer <- reactiveTimer(1000)
  observe({
    copy_timer()
    proc <- rv$copy_proc
    if (is.null(proc)) return(invisible(NULL))
    
    if (!proc$is_alive()) {
      # Just finished
      exit_code <- proc$get_exit_status()
      status <- if (copy_exit_is_success(exit_code)) "SUCCESS" else "FAILED"
      add_log(sprintf("Copy finished: %s (exit code %s). Full log: %s",
                      status, exit_code, rv$copy_log_path))
      rv$copy_proc <- NULL
    }
  })
  
  # -------------------------------------------------------------------------
  # Tab 2: Path Replace
  # -------------------------------------------------------------------------
  
  observeEvent(input$prepare_btn, {
    exts <- trimws(strsplit(input$file_extensions, ",")[[1]])
    root <- effective_dest_root()
    files <- list_files_by_ext(root, exts)
    add_log(sprintf("Scanning %d files under %s for hardcoded path '%s'", length(files), root, hardcoded_path()))
    
    if (length(files) == 0) {
      rv$scan_result <- data.frame(file = character(0), fwd_count = integer(0), bwd_count = integer(0), total = integer(0))
      rv$replace_result <- NULL
      showNotification("No files found to prepare -- copy the folder first.", type = "warning")
      return(invisible(NULL))
    }
    
    pat <- path_patterns(extract_drive_prefix(hardcoded_path()))
    
    # Scan first, so the combined table always shows the reference counts...
    scan_res <- scan_for_drive_refs(files, pat, ignore_case = input$ignore_case)
    scan_res$file <- strip_prefix_literal(scan_res$file, root)
    rv$scan_result <- scan_res[order(-scan_res$total), ]
    add_log(sprintf("Scan complete. %d file(s) contain at least one match.", sum(scan_res$total > 0, na.rm = TRUE)))
    
    # ...then apply the replacement, appending n_replaced/status onto it.
    if (input$keep_backup) {
      backup_dir <- paste0(root, "_backup_", format(Sys.time(), "%Y%m%d_%H%M%S"))
      add_log(paste0("Creating backup copy at: ", backup_dir))
      ok <- tryCatch({
        dir.create(dirname(backup_dir), recursive = TRUE, showWarnings = FALSE)
        file.copy(root, backup_dir, recursive = TRUE)
        TRUE
      }, error = function(e) { add_log(paste0("Backup failed: ", conditionMessage(e))); FALSE })
      if (!ok) {
        showNotification("Backup failed -- preparation aborted so nothing is lost.", type = "error")
        return(invisible(NULL))
      }
    }
    
    add_log(sprintf("Applying replacement '%s' -> '%s' on %d files...", hardcoded_path(), DEFAULT_DEST_ROOT, length(files)))
    apply_res <- apply_drive_replace(files, pat, DEFAULT_DEST_ROOT, ignore_case = input$ignore_case)
    apply_res$file <- strip_prefix_literal(apply_res$file, root)
    rv$replace_result <- apply_res
    n_replaced <- sum(apply_res$status == "REPLACED")
    n_failed <- sum(apply_res$status %in% c("READ_FAILED", "WRITE_FAILED"))
    add_log(sprintf("Files ready for testing: %d modified, %d failed, %d unchanged.",
                    n_replaced, n_failed, sum(apply_res$status == "NO_CHANGE")))
    if (n_failed > 0) {
      showNotification(sprintf("%d file(s) could not be read/written -- check the Console Log.", n_failed), type = "warning")
    }
  })
  
  output$scan_table <- renderDT({
    req(rv$scan_result)
    df <- rv$scan_result
    rep_df <- rv$replace_result
    if (!is.null(rep_df) && nrow(rep_df) > 0) {
      idx <- match(df$file, rep_df$file)
      df$n_replaced <- rep_df$n_replaced[idx]
      df$status <- rep_df$status[idx]
    }
    datatable(df, options = list(pageLength = 10), rownames = FALSE)
  })
  
  # -------------------------------------------------------------------------
  # Tab 3: Run Programs
  # -------------------------------------------------------------------------
  
  refresh_programs <- function() {
    dir_path <- file.path(effective_dest_root(), PROGRAMS_SUBDIR)
    files <- list_files_by_ext(dir_path, c("R"), recursive = isTRUE(input$recurse_subfolders))
    rv$programs <- files
    loa_data(NULL)  # a fresh file listing invalidates any prior LOA-derived order/selection
    add_log(sprintf("Found %d program(s) under %s%s", length(files), dir_path,
                    if (isTRUE(input$recurse_subfolders)) " (including subfolders)" else " (this folder only)"))
  }
  
  observeEvent(input$refresh_programs_btn, {
    refresh_programs()
    load_loa_and_match()
  })
  observeEvent(hardcoded_path(), { refresh_programs() }, ignoreInit = TRUE)
  
  output$programs_table <- renderDT({
    d <- loa_data()
    to_be_run <- if (!is.null(d)) rv$programs %in% d$ordered else rep(FALSE, length(rv$programs))
    df <- data.frame(
      program = basename(rv$programs),
      `To be run` = ifelse(to_be_run, "\u2713", ""),
      check.names = FALSE, stringsAsFactors = FALSE
    )
    selected_idx <- if (!is.null(d)) which(to_be_run) else seq_len(nrow(df))
    datatable(df, selection = list(mode = "multiple", selected = selected_idx),
              options = list(pageLength = 15), rownames = FALSE)
  })
  
  loa_data <- reactiveVal(NULL)
  
  load_loa_and_match <- function() {
    req(input$loa_path)
    if (!file.exists(input$loa_path)) {
      showNotification("LOA file not found at that path.", type = "error")
      add_log(paste0("LOA load failed: file not found: ", input$loa_path))
      return(invisible(NULL))
    }
    
    res <- tryCatch(
      read_loa_yy_programs(input$loa_path, input$loa_status_col, input$loa_program_col, input$loa_sheet_name),
      error = function(e) e
    )
    if (inherits(res, "error")) {
      showNotification(paste0("LOA read error: ", conditionMessage(res)), type = "error")
      add_log(paste0("LOA load failed: ", conditionMessage(res)))
      return(invisible(NULL))
    }
    add_log(sprintf(paste0("LOA loaded (sheet '%s'): %d total row(s), status column '%s', ",
                           "program column '%s', %d program(s) flagged Y/Y"),
                    ifelse(is.na(res$sheet_used), "(n/a -- not Excel)", res$sheet_used),
                    res$n_total_rows, res$status_col, res$program_col, length(res$programs)))
    
    match_res <- match_loa_to_files(res$programs, rv$programs)
    ordered <- order_tf_first(match_res$matched)
    
    # Reorder the program list itself so the to-be-run programs (already
    # _tf-first sorted) show up at the top of the table, in run order --
    # not just a same-order table with some rows ticked.
    rest <- rv$programs[!(rv$programs %in% ordered)]
    rv$programs <- c(ordered, rest)
    
    loa_data(list(ordered = ordered, unmatched = match_res$unmatched_loa_names, n_yy = length(res$programs)))
    
    add_log(sprintf("Matched %d of %d Y/Y program(s) to actual files (%d unmatched).",
                    length(ordered), length(res$programs), length(match_res$unmatched_loa_names)))
    if (length(match_res$unmatched_loa_names) > 0) {
      add_log(paste0("  Not found among discovered files: ", paste(match_res$unmatched_loa_names, collapse = ", ")))
    }
  }
  
  output$programs_select_heading <- renderText({
    n_sel <- length(input$programs_table_rows_selected)
    n_total <- length(rv$programs)
    sprintf("Select programs to run (%d of %d selected)", n_sel, n_total)
  })
  
  output$loa_unmatched_box <- renderUI({
    d <- loa_data()
    if (is.null(d) || length(d$unmatched) == 0) return(NULL)
    div(style = paste0("background:#fff3cd; border:1px solid #ffe69c; ",
                       "border-radius:6px; padding:10px 12px; margin:10px 0;"),
        tags$strong(style = "color:#664d03;",
                    sprintf("%d LOA program(s) not found in this folder (check naming):", length(d$unmatched))),
        tags$ul(style = "margin-bottom:0; padding-left:18px; color:#664d03;",
                lapply(d$unmatched, function(x) tags$li(style = "font-size:0.85rem;", x))
        )
    )
  })
  
  observeEvent(input$run_btn, {
    d <- loa_data()
    using_loa <- !is.null(d) && length(d$ordered) > 0
    if (using_loa) {
      to_run <- d$ordered
    } else {
      sel <- input$programs_table_rows_selected
      if (length(sel) == 0) {
        showNotification("Select at least one program to run.", type = "warning")
        return(invisible(NULL))
      }
      to_run <- rv$programs[sel]
    }
    if (!file.exists(input$rscript_path)) {
      showNotification("Rscript path is not valid. Fix it in the sidebar first.", type = "error")
      return(invisible(NULL))
    }
    if (isTRUE(input$clean_library) && is.null(rv$clean_lib_dir)) {
      rv$clean_lib_dir <- file.path(tempdir(), paste0("clean_rlib_", format(Sys.time(), "%Y%m%d_%H%M%S")))
      dir.create(rv$clean_lib_dir, recursive = TRUE, showWarnings = FALSE)
      add_log(paste0("Clean library created: ", rv$clean_lib_dir,
                     " (packages installed during this run will land here, not in the usual site library)"))
    }
    rv$run_queue <- to_run
    rv$run_results <- list()
    rv$run_active <- TRUE
    add_log(sprintf("Queued %d program(s) to run%s%s.", length(rv$run_queue),
                    if (isTRUE(input$clean_library)) " with a CLEAN package library" else "",
                    if (using_loa) " (LOA Y/Y filter)" else ""))
  })
  
  observeEvent(input$reset_clean_lib, {
    if (!is.null(rv$clean_lib_dir) && dir.exists(rv$clean_lib_dir)) {
      unlink(rv$clean_lib_dir, recursive = TRUE)
    }
    rv$clean_lib_dir <- NULL
    add_log("Clean library reset -- the next run will start with zero pre-installed packages again.")
    showNotification("Clean library wiped.", type = "message")
  })
  
  observeEvent(input$stop_btn, {
    if (!is.null(rv$run_current_proc) && rv$run_current_proc$is_alive()) {
      rv$run_current_proc$kill()
      add_log(paste0("Killed running program: ", basename(rv$run_current_program)))
    }
    rv$run_queue <- character(0)
    rv$run_active <- FALSE
    add_log("Run queue cleared by user (Stop pressed).")
  })
  
  run_timer <- reactiveTimer(700)
  observe({
    run_timer()
    if (!isTRUE(rv$run_active)) return(invisible(NULL))
    
    # Case 1: a program is currently running -- check if it finished
    if (!is.null(rv$run_current_proc)) {
      proc <- rv$run_current_proc
      if (!proc$is_alive()) {
        exit_code <- proc$get_exit_status()
        elapsed <- as.numeric(difftime(Sys.time(), rv$run_current_start, units = "secs"))
        status <- if (!is.na(exit_code) && exit_code == 0) "SUCCESS" else "FAILED"
        log_txt <- tryCatch(paste(readLines(rv$run_current_log, warn = FALSE), collapse = "\n"),
                            error = function(e) "")
        rv$run_results[[basename(rv$run_current_program)]] <- list(
          program = basename(rv$run_current_program),
          exit_code = exit_code,
          status = status,
          seconds = round(elapsed, 1),
          start_time = rv$run_current_start,
          log = log_txt,
          log_file = rv$run_current_log
        )
        add_log(sprintf("Finished: %s -- %s (exit=%s, %.1fs)",
                        basename(rv$run_current_program), status, exit_code, elapsed))
        rv$run_current_proc <- NULL
        rv$run_current_program <- NULL
      } else {
        return(invisible(NULL))  # still running, wait for next tick
      }
    }
    
    # Case 2: nothing running -- start the next queued program, if any
    if (length(rv$run_queue) > 0) {
      next_prog <- rv$run_queue[1]
      rv$run_queue <- rv$run_queue[-1]
      log_path <- file.path(tempdir(), paste0(tools::file_path_sans_ext(basename(next_prog)), "_",
                                              format(Sys.time(), "%Y%m%d_%H%M%S"), ".log"))
      cwd <- dirname(next_prog)
      add_log(sprintf("Starting: %s", basename(next_prog)))
      launch <- build_launch_args(next_prog, input$clean_library, rv$clean_lib_dir)
      proc <- tryCatch({
        processx::process$new(
          input$rscript_path,
          # NOTE: do NOT shQuote() here -- processx passes args directly to the
          # OS process (no shell involved), so shell-quoting would inject
          # literal quote characters into the path and break it.
          args = c(launch$flags, launch$script),
          wd = cwd,
          stdout = log_path,
          stderr = log_path
        )
      }, error = function(e) {
        add_log(sprintf("ERROR launching %s: %s", basename(next_prog), conditionMessage(e)))
        NULL
      })
      rv$run_current_proc <- proc
      rv$run_current_program <- next_prog
      rv$run_current_start <- Sys.time()
      rv$run_current_log <- log_path
    } else {
      if (isTRUE(rv$run_active)) {
        add_log("All queued programs finished.")
      }
      rv$run_active <- FALSE
    }
  })
  
  output$run_status_table <- renderDT({
    if (length(rv$run_results) == 0) {
      return(datatable(data.frame(program = character(0), status = character(0),
                                  exit_code = character(0), seconds = numeric(0)),
                       rownames = FALSE))
    }
    df <- do.call(rbind, lapply(rv$run_results, function(x) {
      data.frame(program = x$program, status = x$status, exit_code = x$exit_code,
                 seconds = x$seconds, stringsAsFactors = FALSE)
    }))
    datatable(df, options = list(pageLength = 15), rownames = FALSE)
  })
  
  # -------------------------------------------------------------------------
  # Tab 4: Output Check
  # -------------------------------------------------------------------------
  
  observeEvent(input$check_btn, {
    output_dir <- file.path(effective_dest_root(), OUTPUT_SUBDIR)
    out_exts <- OUTPUT_EXTENSIONS
    
    programs_ran <- if (length(rv$run_results) > 0) {
      lapply(rv$run_results, function(x) x)
    } else {
      # If nothing was run in this session yet, fall back to all discovered programs
      lapply(rv$programs, function(p) list(program = basename(p), status = "NOT_RUN_THIS_SESSION",
                                           exit_code = NA, seconds = NA, start_time = NA))
    }
    
    rows <- lapply(programs_ran, function(r) {
      prog_name <- r$program
      run_start <- r$start_time
      expected <- guess_expected_outputs(prog_name, output_dir, out_exts)
      
      found <- expected[file.exists(expected)]
      exists_ok <- length(found) > 0
      nonzero_ok <- if (exists_ok) all(file.info(found)$size > 0) else FALSE
      fresh_ok <- if (exists_ok && !is.null(run_start) && !is.na(run_start)) {
        all(file.info(found)$mtime >= run_start)
      } else if (exists_ok) {
        NA
      } else {
        FALSE
      }
      
      data.frame(
        program = prog_name,
        run_status = r$status,
        expected_output = paste(basename(expected), collapse = " | "),
        output_found = exists_ok,
        output_nonzero = nonzero_ok,
        output_fresh = fresh_ok,
        stringsAsFactors = FALSE
      )
    })
    
    rv$completeness_result <- do.call(rbind, rows)
    n_ok <- sum(rv$completeness_result$output_found & rv$completeness_result$output_nonzero, na.rm = TRUE)
    add_log(sprintf("Output check complete: %d/%d program(s) have a non-empty expected output present.",
                    n_ok, nrow(rv$completeness_result)))
  })
  
  output$completeness_table <- renderDT({
    req(rv$completeness_result)
    datatable(rv$completeness_result, options = list(pageLength = 15), rownames = FALSE)
  })
  
  output$save_report_btn <- downloadHandler(
    filename = function() paste0("submission_test_report_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".csv"),
    content = function(file) {
      req(rv$completeness_result)
      write.csv(rv$completeness_result, file, row.names = FALSE)
    }
  )
}

shinyApp(ui, server)