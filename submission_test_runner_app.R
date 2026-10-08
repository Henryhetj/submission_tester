#############################################################################
# SAFE -- Submission Assessment For Executables (Shiny app)
#
# Purpose:
#   Provides a local pre-check of submission executable programs before
#   FDA review: sets up a local working copy of the submission files, runs
#   the programs against it, and checks that the expected outputs were
#   produced.
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
# Dedicated package library for the Windows run, at the root of the user's drive.
WIN_LIB_DIRNAME <- "safe_r_lib"
OUTPUT_SUBDIR <- "output"
OUTPUT_EXTENSIONS <- c("rtf", "docx", "svg")

# Normalize a path prefix (works for both a bare drive letter like "C:" and
# a full path like "C:/submission_files"): trim whitespace and any trailing
# slash/backslash, so we can reliably build "prefix/" and "prefix\" forms.
normalize_prefix <- function(p) {
  p <- trimws(p)
  sub("[/\\\\]+$", "", p)
}

# Convert a path as seen from THIS (Linux) session into the equivalent
# Windows path, for use inside a generated .bat file that will actually run
# on the Windows side. Relies on the shared-storage mapping confirmed
# earlier (this session's home directory and the person's mapped Windows
# drive point at the same underlying storage) -- strips the Linux home
# prefix and re-roots the remaining relative structure under the Windows
# drive letter, converting slash direction along the way.
to_windows_path <- function(linux_path, linux_home, windows_drive) {
  linux_home <- normalize_prefix(linux_home)
  windows_drive <- normalize_prefix(windows_drive)
  rel <- if (startsWith(linux_path, linux_home)) {
    substring(linux_path, nchar(linux_home) + 1)
  } else {
    linux_path
  }
  rel <- gsub("/", "\\\\", rel)
  rel <- sub("^\\\\+", "", rel)
  paste0(windows_drive, "\\", rel)
}

# Strip a literal prefix from the start of a string, for display purposes.
# Deliberately NOT regex-based (no escaping needed, no risk of the prefix
# containing regex metacharacters that break pattern compilation) -- just a
# plain substring check.
strip_prefix_literal <- function(x, prefix) {
  ifelse(startsWith(x, prefix), substring(x, nchar(prefix) + 1), x)
}

# Rewrite any occurrence of the local test root back to the original
# hardcoded path, in free-form text (e.g. captured program output/errors).
# Deliberately a plain literal string replacement (fixed = TRUE), not
# regex-based -- program output is arbitrary text we don't control, and a
# regex approach risks the exact "invalid regex" class of bug fixed earlier
# if the path ever contains characters like parentheses.
restore_hardcoded_path_in_text <- function(text, test_root, hardcoded_root) {
  if (!nzchar(test_root) || is.na(text)) return(text)
  gsub(test_root, hardcoded_root, text, fixed = TRUE)
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

# Replace any hardcoded reference to THIS session's own (Linux) home
# directory with the equivalent Windows drive-letter path -- needed
# because a program (or something it source()s, like autoexec.R) can
# hardcode a path built for the Linux test run specifically (separate from
# the standard "C:/..." convention apply_drive_replace() already handles),
# which is meaningless when the same file is handed to a Windows R engine
# instead. Uses forward slashes in the replacement, since that's valid in
# R source code on Windows too and needs no backslash-escaping the way a
# literal backslash would inside a quoted R string.
apply_linux_home_to_windows_replace <- function(files, linux_home, windows_drive) {
  linux_home <- normalize_prefix(linux_home)
  windows_drive <- normalize_prefix(windows_drive)
  pattern <- paste0(linux_home, "/")
  replacement <- paste0(windows_drive, "/")
  
  out <- lapply(files, function(f) {
    txt <- read_file_text(f)
    if (is.na(txt)) {
      return(data.frame(file = f, n_replaced = NA, status = "READ_FAILED", stringsAsFactors = FALSE))
    }
    orig <- txt
    n_matches <- count_occurrences(txt, pattern)
    txt <- gsub(pattern, replacement, txt, fixed = TRUE)
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

# Compare each local .rtf/.docx output against a same-named file in
# server_dir (top level only -- no subfolder search). Returns a single
# pipe-separated summary string, or NA if there's nothing to compare.
# Neutralize absolute file paths (Windows or Unix style) in extracted text
# before comparing -- footnotes that record "Program Location: /home/.../
# programs/foo.R" etc. will always differ between the local test copy and
# the server original, even when the actual content is identical, since the
# two environments live at different paths.
normalize_paths_for_compare <- function(text) {
  # Restricted to actual path characters (letters, digits, underscore, dot,
  # hyphen, forward slash) -- deliberately excludes backslash, since every
  # hardcoded path seen in these programs uses forward slashes even for the
  # "C:" drive (e.g. "C:/submission_files/..."), and a looser character
  # class would swallow trailing RTF control words (e.g. "\par") that
  # happen to follow a path with no space in between.
  text <- gsub("[A-Za-z]:/[A-Za-z0-9_./-]*", "<PATH>", text, perl = TRUE)
  text <- gsub("/[A-Za-z0-9_.-]+(?:/[A-Za-z0-9_.-]+)+", "<PATH>", text, perl = TRUE)
  text
}

# Extract a comparable text representation of an .rtf or .docx file.
# - .rtf is already largely plain text (with RTF control words mixed in),
#   so it can be read directly.
# - .docx is a zip archive; word/document.xml holds the visible text, so
#   that's extracted and treated as text. This deliberately avoids adding a
#   dependency on a document-parsing package (e.g. officer) just for a
#   same-file-or-not comparison.
# Returns NA on any failure, so the caller can fall back to a raw MD5
# comparison instead.
extract_comparable_text <- function(path) {
  ext <- tolower(tools::file_ext(path))
  if (ext == "docx") {
    tmp_dir <- tempfile("docxcmp")
    dir.create(tmp_dir)
    on.exit(unlink(tmp_dir, recursive = TRUE), add = TRUE)
    ok <- tryCatch({
      utils::unzip(path, files = "word/document.xml", exdir = tmp_dir)
      TRUE
    }, error = function(e) FALSE, warning = function(w) FALSE)
    xml_path <- file.path(tmp_dir, "word", "document.xml")
    if (!ok || !file.exists(xml_path)) return(NA_character_)
    tryCatch(paste(readLines(xml_path, warn = FALSE, encoding = "UTF-8"), collapse = "\n"),
             error = function(e) NA_character_)
  } else {
    tryCatch(paste(readLines(path, warn = FALSE, encoding = "UTF-8"), collapse = "\n"),
             error = function(e) NA_character_)
  }
}

# Find the first point where two strings diverge and return a short,
# human-scannable snippet of context around it -- much more actionable in a
# results table than a bare "DIFFERENT", without dumping the whole file.
find_diff_snippet <- function(text1, text2, context = 15) {
  if (identical(text1, text2)) return("")
  n <- min(nchar(text1), nchar(text2))
  if (n == 0) return("one file is empty")
  c1 <- utf8ToInt(substr(text1, 1, n))
  c2 <- utf8ToInt(substr(text2, 1, n))
  diffs <- which(c1 != c2)
  diff_pos <- if (length(diffs) > 0) diffs[1] else n + 1
  start <- max(1, diff_pos - context)
  snippet1 <- trimws(substr(text1, start, diff_pos + context))
  snippet2 <- trimws(substr(text2, start, diff_pos + context))
  snippet1 <- gsub("[[:space:]]+", " ", snippet1)
  snippet2 <- gsub("[[:space:]]+", " ", snippet2)
  sprintf("local '...%s...' vs server '...%s...'", snippet1, snippet2)
}

compare_outputs_to_server <- function(local_files, server_dir) {
  targets <- local_files[grepl("\\.(rtf|docx)$", local_files, ignore.case = TRUE)]
  if (length(targets) == 0) return(NA_character_)
  if (is.null(server_dir) || !nzchar(trimws(server_dir)) || !dir.exists(server_dir)) {
    return("SERVER_PATH_NOT_FOUND")
  }
  results <- vapply(targets, function(f) {
    ext_tag <- toupper(tools::file_ext(f))
    server_file <- file.path(server_dir, basename(f))
    if (!file.exists(server_file)) return(paste0(ext_tag, ": SERVER_MISSING"))
    
    local_txt <- extract_comparable_text(f)
    server_txt <- extract_comparable_text(server_file)
    
    if (is.na(local_txt) || is.na(server_txt)) {
      # Text extraction failed -- fall back to a raw MD5 comparison rather
      # than silently skipping this file.
      local_md5 <- tryCatch(unname(tools::md5sum(f)), error = function(e) NA_character_)
      server_md5 <- tryCatch(unname(tools::md5sum(server_file)), error = function(e) NA_character_)
      if (is.na(local_md5) || is.na(server_md5)) return(paste0(ext_tag, ": COMPARE_FAILED"))
      return(if (identical(local_md5, server_md5)) paste0(ext_tag, ": MATCH") else paste0(ext_tag, ": DIFFERENT"))
    }
    
    local_norm <- normalize_paths_for_compare(local_txt)
    server_norm <- normalize_paths_for_compare(server_txt)
    if (identical(local_norm, server_norm)) {
      paste0(ext_tag, ": MATCH")
    } else {
      snippet <- find_diff_snippet(local_norm, server_norm)
      paste0(ext_tag, ": DIFFERENT (", snippet, ")")
    }
  }, character(1))
  paste(results, collapse = " | ")
}

# Compare a program's _ards.csv against the server version. PROGRAM/OUTPUT
# columns are always excluded -- they literally record the local vs. server
# file paths, so they're expected to differ even when everything else
# matches. Numeric columns are compared with a small tolerance so
# floating-point noise from different platforms/BLAS libraries doesn't
# produce false alarms; everything else is compared as exact text.
# Compares an ARDS-style CSV against the server version and returns EVERY
# differing cell as its own row (row/column/local/server), rather than
# stopping at the first difference -- so a downloadable report can itemize
# every place two files disagree, not just the first one found. When there's
# nothing to itemize (files match, one is missing, structure differs, or the
# comparison itself failed), returns a single row with row/column left NA
# and the reason in `status`.
compare_ards_to_server_detail <- function(local_csv, server_csv,
                                          ignore_cols = c("PROGRAM", "OUTPUT"), tol = 1e-6) {
  no_diff_row <- function(status) {
    data.frame(row = NA_integer_, column = NA_character_,
               local_value = NA_character_, server_value = NA_character_,
               status = status, stringsAsFactors = FALSE)
  }
  
  if (!file.exists(local_csv)) return(no_diff_row("LOCAL_MISSING"))
  if (!file.exists(server_csv)) return(no_diff_row("SERVER_MISSING"))
  
  # Everything below is wrapped in one tryCatch -- real-world CSVs can have
  # blank/duplicate column headers (e.g. a trailing comma), unexpected
  # types, or other quirks that would otherwise throw an uncaught error and
  # take down the whole comparison (and the reactive session with it, per
  # the on.exit() lesson learned earlier). Any such failure now reports
  # cleanly as COMPARE_FAILED instead.
  tryCatch({
    local_df <- read.csv(local_csv, stringsAsFactors = FALSE, check.names = FALSE)
    server_df <- read.csv(server_csv, stringsAsFactors = FALSE, check.names = FALSE)
    
    # Blank or duplicate headers (a trailing comma, a hand-edited CSV, etc.)
    # would otherwise make column selection by name ambiguous or throw
    # "undefined columns selected" -- give every column a distinct, valid
    # name before comparing.
    fix_names <- function(nms) {
      nms <- trimws(nms)
      nms[!nzchar(nms)] <- "V"
      make.unique(nms)
    }
    names(local_df) <- fix_names(names(local_df))
    names(server_df) <- fix_names(names(server_df))
    
    drop_ignored <- function(df) {
      keep <- !(toupper(names(df)) %in% toupper(ignore_cols))
      df[, keep, drop = FALSE]
    }
    local_df <- drop_ignored(local_df)
    server_df <- drop_ignored(server_df)
    
    if (ncol(local_df) != ncol(server_df) ||
        !identical(sort(names(local_df)), sort(names(server_df)))) {
      return(no_diff_row(sprintf("DIFFERENT (columns differ: local has %d, server has %d)",
                                 ncol(local_df), ncol(server_df))))
    }
    server_df <- server_df[, names(local_df), drop = FALSE]
    
    if (nrow(local_df) != nrow(server_df)) {
      return(no_diff_row(sprintf("DIFFERENT (row count: local=%d, server=%d)", nrow(local_df), nrow(server_df))))
    }
    if (nrow(local_df) == 0) return(no_diff_row("MATCH"))
    
    # Check the RESULT column first -- that's the actual computed value and
    # the signal that actually matters -- but still collect every OTHER
    # column's differences too, not just RESULT's.
    col_order <- names(local_df)
    is_result <- toupper(trimws(col_order)) == "RESULT"
    col_order <- c(col_order[is_result], col_order[!is_result])
    
    all_diffs <- list()
    for (col in col_order) {
      lv <- local_df[[col]]
      sv <- server_df[[col]]
      lnum <- suppressWarnings(as.numeric(lv))
      snum <- suppressWarnings(as.numeric(sv))
      numeric_col <- !any(is.na(lnum) != is.na(lv)) && !any(is.na(snum) != is.na(sv))
      if (numeric_col) {
        diffs <- which(!( (is.na(lnum) & is.na(snum)) | (abs(lnum - snum) <= tol) ))
      } else {
        # Neutralize embedded absolute paths before comparing text columns
        # -- footnote-style columns (e.g. "FOOTNOTE4: Program Location:
        # C:/...") legitimately differ between local and server just
        # because of where each one was run from, same as the docx/rtf
        # comparison already handles. Not restricted to specific column
        # names, since which column holds this text varies by program.
        lchar <- normalize_paths_for_compare(trimws(as.character(lv)))
        schar <- normalize_paths_for_compare(trimws(as.character(sv)))
        diffs <- which(lchar != schar)
      }
      if (length(diffs) > 0) {
        all_diffs[[col]] <- data.frame(
          row = diffs, column = col,
          local_value = as.character(lv[diffs]),
          server_value = as.character(sv[diffs]),
          status = "DIFFERENT",
          stringsAsFactors = FALSE
        )
      }
    }
    
    if (length(all_diffs) == 0) return(no_diff_row("MATCH"))
    out <- do.call(rbind, all_diffs)
    out[order(out$row), ]
  }, error = function(e) {
    no_diff_row(paste0("COMPARE_FAILED (", conditionMessage(e), ")"))
  })
}

# Thin wrapper around compare_ards_to_server_detail() for the on-screen
# summary table -- one line per program, showing just the first difference
# (RESULT column prioritized). The full itemized list is what the
# downloadable report uses instead of this summary.
compare_ards_to_server <- function(local_csv, server_csv,
                                   ignore_cols = c("PROGRAM", "OUTPUT"), tol = 1e-6) {
  d <- compare_ards_to_server_detail(local_csv, server_csv, ignore_cols, tol)
  if (nrow(d) == 1 && is.na(d$row[1])) return(d$status[1])
  i <- 1
  sprintf("DIFFERENT (row %d, col '%s': local='%s' vs server='%s')",
          d$row[i], d$column[i], d$local_value[i], d$server_value[i])
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

# Achieves the "clean package library" isolation via environment variables
# set BEFORE the R process starts (R_LIBS_USER/R_LIBS_SITE/R_LIBS), instead of
# a .libPaths()-in-a-wrapper-script + source() approach.
#
# Why: wrapping the target program in `source(program_path)` changes how
# on.exit() behaves for any top-level (not-inside-a-function) on.exit() calls
# in the target program -- R attaches such calls to the per-statement eval
# frame source() creates, so they fire as soon as the enclosing block
# finishes, not at the true end of the script. Programs whose own cleanup
# logic (sink()/close() for a log file, commonly) relies on running once at
# the very end can then crash ("invalid connection", "no sink to remove")
# purely because of HOW they were invoked -- identical code, different
# result, depending on source() vs direct execution. Confirmed via a minimal
# repro. Setting R_LIBS_* as process environment variables instead means the
# target script is still the literal top-level entry point of the Rscript
# process (exactly as when run directly or from RStudio), so this class of
# bug can't be triggered by this app at all.
build_launch_args <- function(program_path, use_clean_lib, clean_lib_dir) {
  if (!isTRUE(use_clean_lib)) {
    return(list(flags = character(0), script = program_path, env = "current"))
  }
  nonexistent_site_dir <- file.path(clean_lib_dir, "_no_site_library")
  list(
    flags = "--vanilla",
    script = program_path,
    env = c("current",
            R_LIBS_USER = clean_lib_dir,
            R_LIBS_SITE = nonexistent_site_dir,
            R_LIBS = clean_lib_dir)
  )
}

# ---------------------------------------------------------------------------
# Package bootstrap (for the CLEAN-library run and the Windows run script)
#
# The user is never asked where packages come from:
#   * Company-internal packages (e.g. HPC.R.Utilities_*.tar.gz,
#     TRIALImpute_*.tar.gz) are already shipped inside the submission's
#     programs/ folder -> installed straight from those local files.
#   * Public packages -> pulled from the repo this R session is already
#     configured with (Artifactory-backed per company policy). If the session
#     has none configured, fall back to the constant below.
# ---------------------------------------------------------------------------

# TODO(owner): fill in the real Artifactory CRAN remote repo URL (see the
# Artifactory | Developer Platform Front Door page). Only used as a fallback.
LILLY_REPOS_FALLBACK <- c(CRAN = "https://elilillyco.jfrog.io/artifactory/api/cran/<CRAN-REMOTE-REPO-NAME>")

get_package_repos <- function() {
  r <- getOption("repos")
  if (is.null(r) || length(r) == 0 || any(r %in% c("@CRAN@", ""))) LILLY_REPOS_FALLBACK else r
}

# Self-contained on purpose: it is deparse()d into a standalone Rscript (Linux
# clean-lib run) and into the Windows wrapper, so it must not call any other
# function defined in this app.
bootstrap_packages <- function(programs_dir, lib, repos, scan_files, use_pins = TRUE) {
  dir.create(lib, recursive = TRUE, showWarnings = FALSE)
  .libPaths(c(lib, .libPaths()))
  options(repos = repos)
  
  # An install that was interrupted (window closed, paused, network drop) leaves
  # a 00LOCK folder and half-unpacked temp folders behind. A leftover 00LOCK makes
  # EVERY later install.packages() into this library fail, so the packages stay
  # missing and the programs die with "there is no package called ...".
  # This process is the only one touching the library, so it is safe to clear them.
  subdirs <- list.dirs(lib, recursive = FALSE, full.names = TRUE)
  stale <- subdirs[grepl("^00LOCK", basename(subdirs)) | !file.exists(file.path(subdirs, "DESCRIPTION"))]
  if (length(stale) > 0) {
    cat("[bootstrap] removing leftovers of an interrupted install:", paste(basename(stale), collapse = ", "), "\n")
    unlink(stale, recursive = TRUE, force = TRUE)
  }
  
  have <- function() rownames(installed.packages(lib.loc = .libPaths()))
  builtin <- rownames(installed.packages(priority = c("base", "recommended")))
  nrm <- function(p) normalizePath(p, winslash = "/", mustWork = FALSE)
  
  # Scan ONLY the programs being run, plus any .R file they source() (e.g.
  # setup.R / autoexec.R / helpers.R) -- followed transitively. Every other
  # file in the folder is ignored, so unrelated programs/apps in the same
  # folder don't pull in packages this run doesn't need.
  all_r <- nrm(list.files(programs_dir, pattern = "\\.[Rr]$", recursive = TRUE, full.names = TRUE))
  all_r <- all_r[!grepl("(^|/)renv/", all_r)]
  # env.R is always scanned: the run script executes it for every program, and
  # it lists packages as quoted strings (library_loader(c("dplyr", ...))).
  env_r <- file.path(programs_dir, "env.R")
  queue <- unique(nrm(c(scan_files, if (file.exists(env_r)) env_r)))
  files <- character(0)
  followed <- character(0)
  used <- character(0)
  used_in <- character(0)
  while (length(queue) > 0) {
    f <- queue[1]; queue <- queue[-1]
    if (f %in% files) next
    files <- c(files, f)
    txt <- tryCatch(readLines(f, warn = FALSE), error = function(e) character(0))
    txt <- sub("#.*$", "", txt)
    lit <- txt[!grepl("character.only", txt, fixed = TRUE)]
    m_a <- unlist(regmatches(lit, gregexpr("(library|require)\\(\\s*[\"']?[A-Za-z][A-Za-z0-9.]*", lit)))
    m_b <- unlist(regmatches(txt, gregexpr("(requireNamespace|loadNamespace)\\(\\s*[\"'][A-Za-z][A-Za-z0-9.]*", txt)))
    m <- c(m_a, m_b)
    m2 <- unlist(regmatches(txt, gregexpr("[A-Za-z][A-Za-z0-9.]*:::?[A-Za-z_.]", txt)))
    # Packages given as a quoted vector, e.g. library_loader(c("dplyr", "tidyr"))
    flat <- paste(txt, collapse = "\n")
    lm <- unlist(regmatches(flat, gregexpr("library_loader\\(\\s*c\\([^)]*\\)", flat)))
    m3 <- gsub("[\"']", "", unlist(regmatches(lm, gregexpr("[\"'][A-Za-z][A-Za-z0-9.]*[\"']", lm))))
    pk <- c(sub(".*\\(\\s*[\"']?", "", m), sub(":::?.*$", "", m2), m3)
    used <- c(used, pk)
    used_in <- c(used_in, rep(basename(f), length(pk)))
    src_lines <- grep("source\\(", txt, value = TRUE)
    lits <- unlist(regmatches(src_lines, gregexpr("[\"'][^\"']*\\.[Rr][\"']", src_lines)))
    for (n in unique(basename(gsub("[\"']", "", lits)))) {
      hit <- all_r[tolower(basename(all_r)) == tolower(n)]
      new <- setdiff(hit, c(files, queue))
      if (length(new) > 0) {
        queue <- c(queue, new)
        followed <- c(followed, paste(basename(f), "->", n))
      }
    }
  }
  used <- used[nzchar(used)]
  
  # Company packages shipped next to the programs: only the ones actually used
  local_files <- list.files(programs_dir, pattern = "_.*\\.tar\\.gz$", full.names = TRUE)
  local_names <- sub("_.*$", "", basename(local_files))
  keep <- local_names %in% used
  skipped_local <- local_names[!keep]
  local_files <- local_files[keep]
  local_names <- local_names[keep]
  local_deps <- character(0)
  for (i in seq_along(local_files)) {
    tmp <- tempfile(); dir.create(tmp)
    tryCatch(suppressWarnings(untar(local_files[i], files = paste0(local_names[i], "/DESCRIPTION"), exdir = tmp)),
             error = function(e) NULL)
    d <- file.path(tmp, local_names[i], "DESCRIPTION")
    if (file.exists(d)) {
      dc <- read.dcf(d, fields = c("Depends", "Imports", "LinkingTo"))
      x <- unlist(strsplit(paste(na.omit(as.vector(dc)), collapse = ","), ","))
      x <- trimws(gsub("\\(.*$", "", gsub("\\s+", " ", x)))
      local_deps <- c(local_deps, x[nzchar(x) & x != "R"])
    }
  }
  
  # Company packages shipped as .zip (e.g. dataCompareR.zip): either an
  # already-built Windows package or a source folder -- handled both ways.
  zip_files <- list.files(programs_dir, pattern = "\\.zip$", full.names = TRUE)
  zip_names <- sub("(_[0-9][^/]*)?\\.zip$", "", basename(zip_files))
  zkeep <- zip_names %in% used
  zip_files <- zip_files[zkeep]
  zip_names <- zip_names[zkeep]
  zip_dirs <- rep(NA_character_, length(zip_files))
  for (i in seq_along(zip_files)) {
    tmp <- tempfile(); dir.create(tmp)
    tryCatch(unzip(zip_files[i], exdir = tmp), error = function(e) NULL)
    cand <- c(tmp, list.dirs(tmp, recursive = FALSE, full.names = TRUE))
    cand <- cand[file.exists(file.path(cand, "DESCRIPTION"))]
    if (length(cand) > 0) {
      zip_dirs[i] <- cand[1]
      dc <- read.dcf(file.path(cand[1], "DESCRIPTION"), fields = c("Depends", "Imports", "LinkingTo"))
      x <- unlist(strsplit(paste(na.omit(as.vector(dc)), collapse = ","), ","))
      x <- trimws(gsub("\\(.*$", "", gsub("\\s+", " ", x)))
      local_deps <- c(local_deps, x[nzchar(x) & x != "R"])
    }
  }
  local_names <- c(local_names, zip_names)
  
  need <- setdiff(unique(c(used, local_deps)), c(builtin, local_names, have(), "package"))
  cat(sprintf("[bootstrap] scanned %d file(s): %s\n", length(files), paste(basename(files), collapse = ", ")))
  if (length(followed) > 0) cat("[bootstrap] followed source() calls:", paste(followed, collapse = "; "), "\n")
  cat(sprintf("[bootstrap] %d public package(s) needed directly; company package(s) used: %s%s\n",
              length(need), if (length(local_names)) paste(local_names, collapse = ", ") else "none",
              if (length(skipped_local)) paste0(" (not used, skipped: ", paste(skipped_local, collapse = ", "), ")") else ""))
  for (p in need) {
    src <- unique(used_in[used == p])
    cat(sprintf("[bootstrap]   %s <- %s\n", p,
                if (length(src)) paste(utils::head(src, 3), collapse = ", ") else "dependency of a company package"))
  }
  
  if (length(need) > 0) {
    tryCatch(install.packages(need, lib = lib, repos = repos,
                              dependencies = c("Depends", "Imports", "LinkingTo")),
             error = function(e) cat("[bootstrap] public install error:", conditionMessage(e), "\n"))
  }
  for (i in seq_along(local_files)) {
    if (!(local_names[i] %in% have())) {
      cat("[bootstrap] installing company package from local file:", basename(local_files[i]), "\n")
      tryCatch(install.packages(local_files[i], lib = lib, repos = NULL, type = "source"),
               error = function(e) cat("[bootstrap] failed:", basename(local_files[i]), "-", conditionMessage(e), "\n"))
    }
  }
  
  for (i in seq_along(zip_files)) {
    if (zip_names[i] %in% have()) next
    d <- zip_dirs[i]
    if (is.na(d)) {
      cat("[bootstrap] could not find a package inside", basename(zip_files[i]), "\n")
      next
    }
    cat("[bootstrap] installing company package from local file:", basename(zip_files[i]), "\n")
    tryCatch({
      if (dir.exists(file.path(d, "Meta"))) {
        target <- file.path(lib, zip_names[i])
        dir.create(target, showWarnings = FALSE)
        file.copy(list.files(d, full.names = TRUE, all.files = TRUE, no.. = TRUE), target, recursive = TRUE)
      } else {
        install.packages(d, lib = lib, repos = NULL, type = "source")
      }
    }, error = function(e) cat("[bootstrap] failed:", basename(zip_files[i]), "-", conditionMessage(e), "\n"))
  }
  
  # Match the versions pinned in setup.R (the submission's own record of the
  # validated package versions), for packages the programs use. Only pure-R
  # packages (NeedsCompilation: no) are swapped, since older compiled packages
  # would need a compiler toolchain; those keep the version just installed.
  setup_r <- file.path(programs_dir, "setup.R")
  if (isTRUE(use_pins) && file.exists(setup_r)) {
    sl <- readLines(setup_r, warn = FALSE)
    mm <- regmatches(sl, regexec('^\\s*([A-Za-z][A-Za-z0-9.]*)\\s*=\\s*"([0-9][^"]*)"\\s*,?\\s*$', sl))
    mm <- mm[lengths(mm) == 3]
    pins <- setNames(vapply(mm, function(z) z[3], ""), vapply(mm, function(z) z[2], ""))
    snap <- sub('^.*CRAN_URL\\s*<-\\s*"[^"]*/cran/([0-9-]+)".*$', "\\1",
                grep("CRAN_URL\\s*<-", sl, value = TRUE)[1])
    targets <- setdiff(intersect(names(pins), c(used, local_deps)), c(builtin, local_names))
    cat(sprintf("[bootstrap] setup.R pins %d package version(s); checking the %d used by these programs\n",
                length(pins), length(targets)))
    base1 <- unname(repos[1])
    bases <- unique(c(base1, sub("/latest$", paste0("/", snap), base1),
                      sub("__linux__/[^/]+/", "", base1),
                      sub("__linux__/[^/]+/", "", sub("/latest$", paste0("/", snap), base1))))
    for (p in targets) {
      ver <- pins[[p]]
      cur <- tryCatch(as.character(packageVersion(p, lib.loc = .libPaths())), error = function(e) NA_character_)
      if (!is.na(cur) && package_version(cur) == package_version(ver)) next
      tarball <- file.path(tempdir(), sprintf("%s_%s.tar.gz", p, ver))
      got <- FALSE
      for (b in bases) {
        for (u in c(sprintf("%s/src/contrib/Archive/%s/%s_%s.tar.gz", b, p, p, ver),
                    sprintf("%s/src/contrib/%s_%s.tar.gz", b, p, ver))) {
          got <- tryCatch(suppressWarnings(download.file(u, tarball, mode = "wb", quiet = TRUE)) == 0,
                          error = function(e) FALSE)
          if (got) break
        }
        if (got) break
      }
      if (!got) {
        cat(sprintf("[bootstrap]   %s: pinned %s not downloadable, kept %s\n", p, ver, cur))
        next
      }
      tmp <- tempfile(); dir.create(tmp)
      tryCatch(suppressWarnings(untar(tarball, files = paste0(p, "/DESCRIPTION"), exdir = tmp)), error = function(e) NULL)
      dd <- file.path(tmp, p, "DESCRIPTION")
      nc <- if (file.exists(dd)) read.dcf(dd, fields = "NeedsCompilation")[1, 1] else NA
      if (isTRUE(tolower(nc) == "yes")) {
        cat(sprintf("[bootstrap]   %s: pinned %s needs compiling, kept %s\n", p, ver, cur))
        next
      }
      tryCatch({
        install.packages(tarball, lib = lib, repos = NULL, type = "source")
        cat(sprintf("[bootstrap]   %s: %s -> pinned %s\n", p, cur, ver))
      }, error = function(e) cat(sprintf("[bootstrap]   %s: pin to %s failed (%s)\n", p, ver, conditionMessage(e))))
    }
  }
  
  still_missing <- setdiff(c(need, local_names), have())
  if (length(still_missing) > 0) {
    cat("[bootstrap] STILL MISSING:", paste(still_missing, collapse = ", "), "\n")
  } else {
    cat("[bootstrap] all packages available.\n")
  }
  invisible(still_missing)
}

# Lines of a standalone R script that defines + runs bootstrap_packages().
# lib_expr / dir are R code / path strings already valid for the target machine.
bootstrap_script_lines <- function(programs_dir, lib_expr, repos, scan_files, use_pins = TRUE) {
  c(paste0("bootstrap_packages <- ", paste(deparse(bootstrap_packages, width.cutoff = 500L), collapse = "\n")),
    sprintf("bootstrap_packages(%s, %s, %s, %s, %s)", deparse(programs_dir), lib_expr,
            paste(deparse(repos), collapse = ""), paste(deparse(scan_files), collapse = ""),
            deparse(isTRUE(use_pins))))
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
  base_font = bslib::font_collection(
    "-apple-system", "BlinkMacSystemFont", "Segoe UI", "Roboto", "Helvetica Neue", "Arial", "sans-serif"
  ),
  "font-size-base" = "0.85rem",
  "border-radius" = "0.5rem",
  "card-border-color" = "#e2e8f0",
  "navbar-bg" = "#1d4ed8"
)

ui <- bslib::page_sidebar(
  title = "SAFE: Submission Assessment For Executables",
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
  tags$style(HTML(paste0(
    "/* bslib's own '.bslib-gap-spacing' mechanism already zeroes each ",
    "child's own margin-bottom and uses ONE 'gap' value between siblings ",
    "instead -- so the actual fix is turning that gap down (default is a ",
    "generous 1.5rem), not re-adding per-element margins, which would only ",
    "stack on top of the gap and make spacing worse, not better. */",
    ".bslib-gap-spacing { gap: 0.4rem !important; }",
    "h6 { margin-top: 0.6rem; margin-bottom: 0.3rem; font-weight: 600; }",
    "hr { margin: 0.6rem 0 !important; }",
    ".card-header { font-weight: 600; padding-top: 0.6rem; padding-bottom: 0.6rem; }",
    ".help-block { line-height: 1.35 !important; }"
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
            "S"),
        div(
          div(style = "font-weight:700; font-size:0.98rem; line-height:1.25; color:#0f1f3d;",
              "SAFE"),
          div(style = "font-size:0.74rem; color:#334155;",
              "Submission Assessment For Executables")
        )
      ),
      tags$div(style = paste0("font-size:0.66rem; letter-spacing:0.06em; ",
                              "color:#64748b; text-transform:uppercase; ",
                              "margin-bottom:4px;"),
               "About"),
      p(style = "font-size:0.78rem; color:#1e293b; line-height:1.4; margin-bottom:0;",
        "Emulates the FDA reviewer's local machine to test submission ",
        "executable programs before review.")
    ),
    
    tags$div(style = paste0("font-size:0.68rem; letter-spacing:0.06em; ",
                            "opacity:0.55; text-transform:uppercase; ",
                            "margin-bottom:6px;"),
             "Setup"),
    
    h6("1. Submission Program Folder"),
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
    
    h6("4. Server output folder"),
    textInput("server_output_path", NULL,
              value = "/lillyce/qa/ly3437943/j1i_mc_gzbk/final/output/shared"),
    helpText("Compares each program's local .rtf/.docx output against the same-named file here."),
    
    bslib::accordion(
      open = FALSE,
      bslib::accordion_panel(
        "Advanced settings",
        h6("R engine"),
        radioButtons("r_engine_mode", NULL,
                     choices = c("Server R engine (Linux, default)" = "default",
                                 "Local R engine (Windows)" = "custom"),
                     selected = "default"),
        conditionalPanel(
          condition = "input.r_engine_mode == 'custom'",
          div(style = paste0("background:#fff3cd; border:1px solid #ffe69c; ",
                             "border-radius:6px; padding:8px 12px; margin-bottom:8px; font-size:0.8rem; color:#664d03;"),
              tags$strong("R must already be installed locally "),
              "for this to work -- this only points at an existing install; it doesn't install R for you. ",
              "This session can't launch a Windows program directly either, so 'Run selected programs' will ",
              "instead generate a .R script for you to run with one typed command on the Windows side. ",
              "It may also adjust files in the shared 'programs' folder to work on Windows -- re-run ",
              "'Set up testing environment' before switching back to the Linux engine."),
          textInput("rscript_path", "Rscript path", value = "Y:/Programs/R/R-4.6.1/bin/Rscript.exe"),
          helpText("Packages for this run install into <drive>:\\", WIN_LIB_DIRNAME, " (drive taken from the Rscript path), ",
                   "so nothing goes to C: or the R install folder. 'Reset clean library' below deletes it.")
        ),
        helpText("Use your own local R install (e.g. on your Citrix Y: drive) instead of the server R engine. ",
                 "Point at the Rscript.exe inside that install's bin/ folder, not RStudio.exe."),
        checkboxInput("clean_library", "Run with a CLEAN package library (exclude pre-installed extension packages)",
                      value = TRUE),
        helpText("Forces each program to install its own packages instead of using pre-installed ones. ",
                 "Persists within one Run, resets on the next -- still not a substitute for the Windows/Citrix toolchain."),
        checkboxInput("use_pinned_versions", "Use the package versions pinned in setup.R where possible (slower)", value = FALSE),
        helpText("Matches the validated versions listed in the submission's setup.R (e.g. officer, flextable) for ",
                 "packages these programs use. Only pure-R packages are swapped; compiled ones keep the current version. ",
                 "Off by default: packages are only checked for being installed, not for version. ",
                 "Takes effect when the library is (re)built -- use 'Reset clean library' to start over."),
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
        checkboxInput("keep_backup", "Keep an untouched backup before replacing", value = FALSE),
        checkboxInput("force_recopy", "Always re-copy files, even if already set up", value = FALSE)
      )
    )
  ),
  
  bslib::layout_columns(
    col_widths = c(6, 6),
    
    bslib::card(
      full_screen = TRUE,
      height = "780px",
      bslib::card_header("1. Set Up Testing Environment & Run Programs"),
      div(style = "display: flex; flex-direction: row; align-items: center; flex-wrap: wrap; gap: 8px 16px;",
          actionButton("setup_env_btn", "Set up testing environment", class = "btn-primary btn-sm", style = "width: auto;"),
          actionButton("refresh_programs_btn", "Refresh program list based on LOA", class = "btn-primary btn-sm", style = "width: auto;"),
          div(style = "white-space: nowrap;", checkboxInput("recurse_subfolders", "Include subfolders", value = FALSE))
      ),
      helpText("Finds which programs are ready to test, based on your LOA file."),
      h6(textOutput("programs_select_heading", inline = TRUE)),
      div(style = "display: flex; flex-direction: row; gap: 8px; margin-bottom: 4px;",
          actionButton("select_all_btn", "Select all", class = "btn-outline-primary btn-sm", style = "width: auto;"),
          actionButton("deselect_all_btn", "Deselect all", class = "btn-outline-secondary btn-sm", style = "width: auto;")
      ),
      DTOutput("programs_table"),
      uiOutput("loa_unmatched_box"),
      div(style = "display: flex; flex-direction: row; flex-wrap: wrap; gap: 8px;",
          actionButton("run_btn", "Run selected programs", class = "btn-primary btn-sm", style = "width: auto;"),
          actionButton("stop_btn", "Stop", class = "btn-danger btn-sm", style = "width: auto;")
      ),
      h6("Run status:"),
      DTOutput("run_status_table")
    ),
    
    bslib::card(
      full_screen = TRUE,
      height = "780px",
      bslib::card_header("2. Output & Log"),
      h6("Output Check"),
      div(style = "display: flex; flex-direction: row; flex-wrap: wrap; gap: 8px;",
          actionButton("check_btn", "Check output completeness", class = "btn-primary btn-sm", style = "width: auto;"),
          downloadButton("save_report_btn", "Download report (CSV)", class = "btn-sm", style = "width: auto;")
      ),
      DTOutput("completeness_table"),
      
      hr(),
      h6("Output Comparison"),
      div(style = "display: flex; flex-direction: row; flex-wrap: wrap; gap: 8px;",
          actionButton("compare_ards_btn", "Compare output to server", class = "btn-primary btn-sm", style = "width: auto;"),
          downloadButton("save_ards_report_btn", "Download report (CSV)", class = "btn-sm", style = "width: auto;")
      ),
      helpText("Compares each LOA-required program's output from Step 1's folder against the server version."),
      DTOutput("ards_compare_table"),
      
      hr(),
      actionButton("clear_cache_btn", "Clear testing environment", class = "btn-danger btn-sm", style = "width: auto;"),
      helpText("Deletes the local test copy so the next setup starts fresh."),
      
      hr(),
      div(style = "display: flex; flex-direction: row; justify-content: space-between; align-items: center;",
          h6("Console Log", style = "margin-bottom: 0;"),
          actionButton("clear_log_btn", "Clear log", class = "btn-sm", style = "width: auto;")
      ),
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
    bootstrap_proc = NULL,
    bootstrap_log = NULL,
    completeness_result = NULL,
    ards_compare_result = NULL,
    ards_compare_detail = NULL,
    auto_chain_pending = FALSE
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
  
  # Which R engine actually runs the programs: this app's own (whatever R
  # process is hosting the app itself -- e.g. the Linux Posit Workbench
  # engine, even when accessed through a Windows/Citrix browser), or a
  # specific alternate install the person points at (e.g. one they've
  # installed locally on their Citrix Y: drive), to test against a
  # genuinely different R engine/package setup.
  effective_rscript_path <- reactive({
    if (identical(input$r_engine_mode, "custom")) input$rscript_path else DEFAULT_RSCRIPT_PATH
  })
  
  observeEvent(input$clear_cache_btn, {
    dst <- effective_dest_root()
    if (dir.exists(dst)) {
      unlink(dst, recursive = TRUE)
      add_log("Testing environment cleared.")
      showNotification("Testing environment cleared.", type = "message")
    } else {
      add_log("Nothing to clear -- testing environment doesn't exist.")
    }
  })
  
  output$hardcoded_path_breakdown <- renderText({
    sprintf("Detected: %s", protocol_id_reactive())
  })
  
  # -------------------------------------------------------------------------
  # Tab 1: Copy
  # -------------------------------------------------------------------------
  
  run_copy_step <- function() {
    src <- input$copy_from_path
    dst <- effective_dest_root()
    
    if (!dir.exists(src)) {
      add_log(paste0("ERROR: source path does not exist or is not reachable: ", src))
      showNotification("Source path not found. Check the path and network connectivity.", type = "error")
      return(FALSE)
    }
    
    rv$copy_total_files <- length(list.files(src, recursive = TRUE))
    add_log(sprintf("Setting up files (%d found)...", rv$copy_total_files))
    
    log_path <- file.path(tempdir(), paste0("copy_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".log"))
    rv$copy_log_path <- log_path
    
    proc <- tryCatch({
      start_copy_process(src, dst, log_path)
    }, error = function(e) {
      add_log(paste0("ERROR launching copy: ", conditionMessage(e)))
      NULL
    })
    rv$copy_proc <- proc
    !is.null(proc)
  }
  
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
      add_log(if (status == "SUCCESS") "Files ready." else sprintf("File setup failed (exit code %s). Check network/path access.", exit_code))
      rv$copy_proc <- NULL
      
      # If this copy was kicked off by "Set up testing environment" (Step
      # 1), silently continue with preparing the files once the copy
      # finishes -- the person clicked one button for "set up", not two.
      if (isTRUE(rv$auto_chain_pending)) {
        rv$auto_chain_pending <- FALSE
        if (status == "SUCCESS") {
          run_prepare_step()
        } else {
          showNotification("Setup failed -- see Console Log for details.", type = "error")
        }
      }
    }
  })
  
  # -------------------------------------------------------------------------
  # Tab 2: Path Replace
  # -------------------------------------------------------------------------
  
  run_prepare_step <- function() {
    exts <- trimws(strsplit(input$file_extensions, ",")[[1]])
    root <- effective_dest_root()
    files <- list_files_by_ext(root, exts)
    add_log(sprintf("Checking %d file(s) for hardcoded path '%s'...", length(files), hardcoded_path()))
    
    if (length(files) == 0) {
      rv$scan_result <- data.frame(file = character(0), fwd_count = integer(0), bwd_count = integer(0), total = integer(0))
      rv$replace_result <- NULL
      showNotification("No files found to prepare.", type = "warning")
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
      add_log("Creating a backup copy before making changes...")
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
    
    add_log(sprintf("Updating %d file(s) that reference '%s'...", length(files), hardcoded_path()))
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
    invisible(NULL)
  }
  
  # -------------------------------------------------------------------------
  # Tab 3: Run Programs
  # -------------------------------------------------------------------------
  
  refresh_programs <- function() {
    dir_path <- file.path(effective_dest_root(), PROGRAMS_SUBDIR)
    files <- list_files_by_ext(dir_path, c("R"), recursive = isTRUE(input$recurse_subfolders))
    rv$programs <- files
    loa_data(NULL)  # a fresh file listing invalidates any prior LOA-derived order/selection
    add_log(sprintf("Found %d program(s)%s", length(files),
                    if (isTRUE(input$recurse_subfolders)) " (including subfolders)" else ""))
  }
  
  observeEvent(input$setup_env_btn, {
    # Step 1: copy -> prepare. Skip the copy step entirely if the
    # destination already has files from a previous run -- copying is
    # usually the slowest part, and re-copying unchanged files on every
    # click just wastes time (see "Force re-copy" in Advanced settings).
    dst_dir <- effective_dest_root()
    already_present <- !isTRUE(input$force_recopy) && dir.exists(dst_dir) &&
      length(list.files(dst_dir, recursive = TRUE)) > 0
    
    if (already_present) {
      add_log("Files already set up -- skipping to the next step.")
      run_prepare_step()
      return(invisible(NULL))
    }
    
    rv$auto_chain_pending <- TRUE
    ok <- run_copy_step()
    if (!ok) rv$auto_chain_pending <- FALSE
  })
  
  observeEvent(input$refresh_programs_btn, {
    refresh_programs()
    load_loa_and_match()
  })
  observeEvent(hardcoded_path(), { refresh_programs() }, ignoreInit = TRUE)
  
  output$programs_table <- renderDT({
    d <- loa_data()
    to_be_run <- if (!is.null(d)) rv$programs %in% d$ordered else rep(FALSE, length(rv$programs))
    df <- data.frame(program = basename(rv$programs), stringsAsFactors = FALSE)
    # Row highlighting IS "will be run" -- no separate "To be run" column
    # that could get out of sync with it. LOA matches just set the initial
    # selection; the person can then freely check/uncheck rows, and
    # whatever's actually highlighted when they click Run is what runs.
    selected_idx <- if (!is.null(d)) which(to_be_run) else seq_len(nrow(df))
    datatable(df, selection = list(mode = "multiple", selected = selected_idx),
              options = list(pageLength = 15), rownames = FALSE)
  })
  
  programs_proxy <- DT::dataTableProxy("programs_table")
  
  # "Select all" respects the table's Search box: it selects the rows currently
  # shown after filtering (all rows when there is no search text).
  observeEvent(input$select_all_btn, {
    rows <- input$programs_table_rows_all
    if (is.null(rows)) rows <- seq_along(rv$programs)
    DT::selectRows(programs_proxy, rows)
  })
  
  observeEvent(input$deselect_all_btn, {
    DT::selectRows(programs_proxy, NULL)
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
    
    loa_data(list(ordered = ordered, unmatched = match_res$unmatched_loa_names,
                  n_yy = length(res$programs), all_yy = res$programs))
    
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
    # Always run exactly what's currently highlighted in the table --
    # loading the LOA only sets the INITIAL selection (the LOA Y/Y
    # matches); after that, whatever the person has manually checked or
    # unchecked is what actually runs. No special-case bypass that ignores
    # a manual deselection just because LOA data happens to be loaded.
    sel <- input$programs_table_rows_selected
    if (length(sel) == 0) {
      showNotification("Select at least one program to run.", type = "warning")
      return(invisible(NULL))
    }
    to_run <- rv$programs[sel]
    
    if (identical(input$r_engine_mode, "custom")) {
      # This (Linux) session can't launch a Windows process directly, and
      # corporate security policy commonly blocks .bat/.cmd/.ps1 files
      # specifically (even from a trusted-enough network location) while
      # still trusting the Rscript.exe binary itself. So generate a plain
      # .R wrapper script instead -- the person runs it with ONE typed
      # command that invokes the already-trusted Rscript.exe directly on
      # this file, sidestepping the blocked script types entirely.
      #
      # Inside the wrapper, each program is launched as its OWN separate
      # Rscript.exe process via system2() -- NOT source(). Source()-ing one
      # script from inside another changes how a top-level (not-in-a-
      # function) on.exit() behaves in the called script, which is exactly
      # the bug this app's own build_launch_args() was redesigned to avoid
      # earlier. Each program needs to stay the literal top-level entry
      # point of its own process, same as it would running directly.
      linux_home <- Sys.getenv("HOME")
      win_rscript <- input$rscript_path
      # Auto-detect the drive letter from the Rscript path itself (e.g.
      # "Y:/Programs/..." -> "Y:") instead of asking for it separately --
      # it's already right there at the start of the path.
      win_drive <- extract_drive_prefix(win_rscript)
      if (!grepl("^[A-Za-z]:$", win_drive)) {
        showNotification("Rscript path should start with a drive letter (e.g. 'Y:/Programs/...') so it can be detected automatically.", type = "error")
        return(invisible(NULL))
      }
      
      # Some programs (or things they source(), like autoexec.R) can
      # hardcode an absolute reference to THIS session's own home
      # directory -- correct for the Linux test run, meaningless on
      # Windows. Fixed IN PLACE, in the same "programs" folder everything
      # else already uses -- NOT a separately-named copy. Programs in this
      # folder can reference each other (e.g. by a path built from their
      # OWN folder's name), so a differently-named sibling folder risks a
      # subtler, harder-to-spot path bug than the one this is fixing.
      # Trade-off: this does mean switching back to the Linux engine
      # afterward needs "Set up testing environment" run again (ideally
      # with "Force re-copy" on) to get a clean Linux-appropriate copy.
      programs_dir <- file.path(effective_dest_root(), PROGRAMS_SUBDIR)
      all_program_files <- list_files_by_ext(programs_dir, c("R"), recursive = TRUE)
      home_fix <- apply_linux_home_to_windows_replace(all_program_files, linux_home, win_drive)
      n_home_fixed <- sum(home_fix$status == "REPLACED", na.rm = TRUE)
      if (n_home_fixed > 0) {
        add_log(sprintf(paste0("Adjusted %d file(s) that referenced this session's own path, for the Windows run ",
                               "(re-run 'Set up testing environment' before switching back to the Linux engine)."),
                        n_home_fixed))
      }
      
      win_programs <- vapply(to_run, function(p) to_windows_path(p, linux_home, win_drive), character(1))
      
      win_programs_dir <- to_windows_path(programs_dir, linux_home, win_drive)
      
      # Dedicated package library: packages never land in the person's C:
      # user library / the R install itself, and the run can't silently reuse
      # packages installed earlier. Lives at the root of the person's drive
      # (outside the copied submission folder, so the path-replace step never
      # touches installed packages).
      win_lib <- to_windows_path(file.path(normalize_prefix(linux_home), WIN_LIB_DIRNAME), linux_home, win_drive)
      add_log(sprintf("Windows run will use a dedicated package library: %s", win_lib))
      lib_env_lines <- c(
        # Forward slashes on purpose: R reports .libPaths() entries that way, and
        # the TFL programs compare them to Sys.getenv("R_LIBS_USER") as plain text.
        paste0("lib_dir <- ", deparse(gsub("\\\\", "/", win_lib))),
        "Sys.setenv(R_LIBS = lib_dir, R_LIBS_USER = lib_dir,",
        "           R_LIBS_SITE = lib_dir)",
        'cat("Package library for this run:", lib_dir, "\\n\\n")'
      )
      
      # Diagnostic hook: if a program dies with an uncaught error, print which
      # library paths that R process had at that moment (and where dplyr is, if
      # anywhere). Loaded into every Rscript child via R_PROFILE_USER.
      diag_profile_lines <- c(
        "if (nzchar(Sys.getenv('R_DIAG_TRACE'))) {",
        "invisible(tryCatch(suppressMessages(trace('.libPaths', tracer = quote(if (!missing(new)) { cat('\\n[diag] .libPaths() reset to: ', paste(new, collapse = ' | '), '\\n', sep = ''); cat(paste0('  ', vapply(utils::tail(sys.calls(), 8), function(x) substr(paste(deparse(x), collapse = ' '), 1, 160), '')), sep = '\\n') }), print = FALSE)), error = function(e) NULL))",
        "options(error = function() {",
        "  cat('\\n[diag] library paths at the time of the error:\\n')",
        "  cat(paste0('  ', .libPaths()), sep = '\\n')",
        "  cat('[diag] R_LIBS = ', Sys.getenv('R_LIBS'), '\\n', sep = '')",
        "  cat('[diag] working directory = ', getwd(), '\\n', sep = '')",
        "  cat('[diag] dplyr found at: ', system.file(package = 'dplyr'), '\\n', sep = '')",
        "  quit(save = 'no', status = 1)",
        "})",
        "}"
      )
      diag_lines <- c(
        "diag_profile <- file.path(tempdir(), 'diag_profile.R')",
        paste0("writeLines(", paste(deparse(diag_profile_lines), collapse = ""), ", diag_profile)"),
        "Sys.setenv(R_PROFILE_USER = diag_profile)"
      )
      
      # Package preparation runs as its OWN Rscript process (with the same
      # library environment as the programs), so what it sees and installs is
      # exactly what each program will see.
      boot_path <- file.path(effective_dest_root(), "bootstrap_packages.R")
      writeLines(bootstrap_script_lines(win_programs_dir,
                                        deparse(win_lib),
                                        get_package_repos(), win_programs, isTRUE(input$use_pinned_versions)),
                 boot_path)
      win_boot_path <- to_windows_path(boot_path, linux_home, win_drive)
      bootstrap_lines <- c(
        'cat("=== Preparing packages (company packages from programs folder, public from Artifactory) ===\\n")',
        paste0("system2(rscript_exe, args = shQuote(", deparse(win_boot_path), "))"),
        'cat("=== Package preparation done ===\\n\\n")'
      )
      add_log("Including package preparation in the Windows run script (company packages come from the programs folder).")
      
      # Quick sanity check, in a fresh process with the same environment the
      # programs get: which library paths does R see, and can it find dplyr?
      check_lines <- c(
        'cat("=== Checking what the programs will see ===\\n")',
        'system2(rscript_exe, args = c("-e", shQuote("cat(\'Library paths:\', .libPaths(), sep = \'\\\\n  \'); cat(\'\\\\ndplyr available:\', requireNamespace(\'dplyr\', quietly = TRUE), \'\\\\n\\\\n\')")))'
      )
      
      # If the project has an env.R with install-if-missing package setup,
      # run it once too (also as its own process, same library environment).
      env_r_path <- file.path(programs_dir, "env.R")
      env_setup_lines <- if (file.exists(env_r_path)) {
        win_env_r <- to_windows_path(env_r_path, linux_home, win_drive)
        add_log("Including env.R in the Windows run script to install any missing packages first.")
        c(
          'cat("=== Installing/loading required packages (env.R) ===\\n")',
          paste0("system2(rscript_exe, args = shQuote(", deparse(win_env_r), "))"),
          'cat("=== Package setup done ===\\n\\n")'
        )
      } else {
        character(0)
      }
      
      wrapper_lines <- c(
        paste0("rscript_exe <- ", deparse(win_rscript)),
        "programs <- c(",
        paste0("  ", vapply(win_programs, deparse, character(1)),
               c(rep(",", length(win_programs) - 1), "")),
        ")",
        lib_env_lines,
        diag_lines,
        bootstrap_lines,
        check_lines,
        env_setup_lines,
        "Sys.setenv(R_DIAG_TRACE = '1')",
        "for (p in programs) {",
        '  cat("=== Running:", p, "===\\n")',
        "  status <- system2(rscript_exe, args = shQuote(p))",
        '  cat("=== Finished (exit code", status, "):", p, "===\\n\\n")',
        "}",
        'cat("All programs finished. Press Enter to close this window.\\n")',
        "invisible(readline())"
      )
      wrapper_path <- file.path(effective_dest_root(), "run_selected_programs.R")
      writeLines(wrapper_lines, wrapper_path)
      win_wrapper_path <- to_windows_path(wrapper_path, linux_home, win_drive)
      
      # Show the Rscript path with backslashes too, so both halves of the
      # command use the same separator style as the wrapper path above.
      win_rscript_display <- gsub("/", "\\\\", win_rscript)
      run_command <- sprintf('"%s" "%s"', win_rscript_display, win_wrapper_path)
      add_log(sprintf("Generated a Windows run script for %d program(s): %s", length(win_programs), win_wrapper_path))
      add_log(sprintf("Run it by typing this in a Windows Command Prompt: %s", run_command))
      
      # Best-effort: copy the command straight to the clipboard so there's
      # nothing to select -- just switch to the Command Prompt and paste.
      # Some locked-down browser/security setups block clipboard access
      # outright, so this can silently fail -- the command is still shown
      # above (and in the notification below) either way, for manual copy.
      if (requireNamespace("jsonlite", quietly = TRUE)) {
        shinyjs::runjs(sprintf(
          "navigator.clipboard.writeText(%s).catch(function(e) { console.log('Clipboard copy failed:', e); });",
          jsonlite::toJSON(run_command, auto_unbox = TRUE)
        ))
      }
      
      showNotification(sprintf("Script created and command copied to clipboard. Paste it into a Windows Command Prompt: %s", run_command),
                       type = "message", duration = 20)
      return(invisible(NULL))
    }
    
    if (!file.exists(effective_rscript_path())) {
      showNotification("Rscript path is not valid. Fix it in Advanced settings first.", type = "error")
      return(invisible(NULL))
    }
    if (isTRUE(input$clean_library) && is.null(rv$clean_lib_dir)) {
      rv$clean_lib_dir <- file.path(tempdir(), paste0("clean_rlib_", format(Sys.time(), "%Y%m%d_%H%M%S")))
      dir.create(rv$clean_lib_dir, recursive = TRUE, showWarnings = FALSE)
      add_log("Clean package library set up (packages will install fresh during this run).")
    }
    rv$run_queue <- to_run
    rv$run_results <- list()
    add_log(sprintf("Queued %d program(s) to run%s.", length(rv$run_queue),
                    if (isTRUE(input$clean_library)) " with a CLEAN package library" else ""))
    
    if (isTRUE(input$clean_library)) {
      # A clean library has zero packages, so first make everything the
      # programs need available: company packages from the programs folder,
      # public ones from the company-configured repo. No user input needed.
      boot_script <- file.path(tempdir(), "bootstrap_packages.R")
      writeLines(bootstrap_script_lines(file.path(effective_dest_root(), PROGRAMS_SUBDIR),
                                        deparse(rv$clean_lib_dir), get_package_repos(), to_run,
                                        isTRUE(input$use_pinned_versions)),
                 boot_script)
      rv$bootstrap_log <- file.path(tempdir(), paste0("bootstrap_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".log"))
      launch <- build_launch_args(boot_script, TRUE, rv$clean_lib_dir)
      rv$bootstrap_proc <- tryCatch(
        processx::process$new(effective_rscript_path(), args = c(launch$flags, launch$script),
                              env = launch$env, wd = dirname(boot_script),
                              stdout = rv$bootstrap_log, stderr = rv$bootstrap_log),
        error = function(e) { add_log(paste0("ERROR launching package preparation: ", conditionMessage(e))); NULL })
      if (!is.null(rv$bootstrap_proc)) {
        add_log("Preparing packages first (company packages from the programs folder, public ones from the configured repo)...")
        return(invisible(NULL))
      }
    }
    rv$run_active <- TRUE
  })
  
  observeEvent(input$reset_clean_lib, {
    if (!is.null(rv$clean_lib_dir) && dir.exists(rv$clean_lib_dir)) {
      unlink(rv$clean_lib_dir, recursive = TRUE)
    }
    win_lib_linux <- file.path(normalize_prefix(Sys.getenv("HOME")), WIN_LIB_DIRNAME)
    if (dir.exists(win_lib_linux)) {
      unlink(win_lib_linux, recursive = TRUE)
      add_log(sprintf("Removed the dedicated Windows package library (%s).", WIN_LIB_DIRNAME))
    }
    rv$clean_lib_dir <- NULL
    add_log("Clean library reset -- the next run will start with zero pre-installed packages again.")
    showNotification("Clean library wiped.", type = "message")
  })
  
  observeEvent(input$stop_btn, {
    if (!is.null(rv$bootstrap_proc) && rv$bootstrap_proc$is_alive()) {
      rv$bootstrap_proc$kill()
      rv$bootstrap_proc <- NULL
      add_log("Package preparation stopped by user.")
    }
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
    proc <- rv$bootstrap_proc
    if (is.null(proc) || proc$is_alive()) return(invisible(NULL))
    txt <- tryCatch(paste(readLines(rv$bootstrap_log, warn = FALSE), collapse = "\n"), error = function(e) "")
    lines <- grep("^\\[bootstrap\\]", strsplit(txt, "\n", fixed = TRUE)[[1]], value = TRUE)
    if (length(lines) > 0) add_log(paste(lines, collapse = "\n"))
    rv$bootstrap_proc <- NULL
    if (length(rv$run_queue) > 0) rv$run_active <- TRUE
  })
  observe({
    run_timer()
    if (!isTRUE(rv$run_active)) return(invisible(NULL))
    
    # Case 1: a program is currently running -- check if it finished
    if (!is.null(rv$run_current_proc)) {
      proc <- rv$run_current_proc
      if (!proc$is_alive()) {
        tryCatch({
          exit_code <- proc$get_exit_status()
          elapsed <- as.numeric(difftime(Sys.time(), rv$run_current_start, units = "secs"))
          status <- if (!is.na(exit_code) && exit_code == 0) "SUCCESS" else "FAILED"
          prog_name <- basename(rv$run_current_program)
          log_txt <- tryCatch({
            raw <- paste(readLines(rv$run_current_log, warn = FALSE), collapse = "\n")
            # Program output can legitimately contain non-UTF-8 bytes (locale-
            # dependent characters, garbled output from a crashing package,
            # etc.) -- replace invalid byte sequences instead of letting
            # downstream string ops (trimws, sub, ...) throw on them.
            sanitized <- iconv(raw, from = "UTF-8", to = "UTF-8", sub = "byte")
            # Show the ORIGINAL hardcoded path (e.g. "C:/submission_files/...")
            # in place of the local test root, so an error/warning naming a
            # file path reads as a direct, actionable statement about the
            # real path -- not a Linux path the person has to mentally
            # translate back themselves.
            restore_hardcoded_path_in_text(sanitized, effective_dest_root(), hardcoded_path())
          }, error = function(e) "")
          
          rv$run_results[[prog_name]] <- list(
            program = prog_name,
            exit_code = exit_code,
            status = status,
            seconds = round(elapsed, 1),
            start_time = rv$run_current_start,
            log = log_txt,
            log_file = rv$run_current_log
          )
          add_log(sprintf("Finished: %s -- %s (exit=%s, %.1fs)", prog_name, status, exit_code, elapsed))
          
          if (status == "FAILED" && nzchar(trimws(log_txt))) {
            lines <- strsplit(log_txt, "\n", fixed = TRUE)[[1]]
            n_show <- 40
            shown <- if (length(lines) > n_show) utils::tail(lines, n_show) else lines
            note <- if (length(lines) > n_show) sprintf(" (last %d of %d lines)", n_show, length(lines)) else ""
            add_log(sprintf("Error output for %s%s:\n%s", prog_name, note, paste(shown, collapse = "\n")))
          }
        }, error = function(e) {
          # Never let an unexpected error here stall the run loop forever --
          # log it and still let the queue move on (state cleared below,
          # outside this tryCatch, unconditionally).
          add_log(sprintf("Internal error while finishing %s: %s",
                          basename(rv$run_current_program), conditionMessage(e)))
        })
        
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
          effective_rscript_path(),
          # NOTE: do NOT shQuote() here -- processx passes args directly to the
          # OS process (no shell involved), so shell-quoting would inject
          # literal quote characters into the path and break it.
          args = c(launch$flags, launch$script),
          env = launch$env,
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
    source_output_dir <- file.path(normalize_prefix(input$copy_from_path), OUTPUT_SUBDIR)
    out_exts <- OUTPUT_EXTENSIONS
    
    d <- loa_data()
    target_programs <- if (!is.null(d) && length(d$all_yy) > 0) {
      # Every LOA-required program, regardless of whether its .R file was
      # actually found locally -- checking OUTPUT doesn't require the
      # program itself to be present (e.g. outputs copied in directly).
      d$all_yy
    } else if (length(rv$run_results) > 0) {
      names(rv$run_results)
    } else {
      basename(rv$programs)
    }
    
    programs_ran <- lapply(target_programs, function(prog_name) {
      base <- tools::file_path_sans_ext(basename(prog_name))
      # Match against actual run results case/extension-insensitively,
      # since LOA-listed names don't always include ".R" or exact case.
      run_key <- names(rv$run_results)[
        tolower(tools::file_path_sans_ext(names(rv$run_results))) == tolower(base)
      ]
      if (length(run_key) > 0) {
        rv$run_results[[run_key[1]]]
      } else {
        list(program = base, status = "NOT_RUN_THIS_SESSION",
             exit_code = NA, seconds = NA, start_time = NA)
      }
    })
    
    rows <- lapply(programs_ran, function(r) {
      prog_name <- r$program
      run_start <- r$start_time
      
      # Check the local TEST environment's output first; fall back to the
      # original source folder's own output (e.g. outputs copied there
      # directly, without re-running the program locally).
      expected_test <- guess_expected_outputs(prog_name, output_dir, out_exts)
      expected_source <- guess_expected_outputs(prog_name, source_output_dir, out_exts)
      expected <- ifelse(file.exists(expected_test), expected_test, expected_source)
      
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
      
      server_compare <- if (exists_ok) compare_outputs_to_server(found, input$server_output_path) else NA_character_
      
      data.frame(
        program = prog_name,
        run_status = r$status,
        expected_output = paste(basename(expected), collapse = " | "),
        output_found = exists_ok,
        output_nonzero = nonzero_ok,
        output_fresh = fresh_ok,
        server_compare = server_compare,
        stringsAsFactors = FALSE
      )
    })
    
    rv$completeness_result <- do.call(rbind, rows)
    n_ok <- sum(rv$completeness_result$output_found & rv$completeness_result$output_nonzero, na.rm = TRUE)
    add_log(sprintf("Output check complete: %d/%d program(s) have a non-empty expected output present.",
                    n_ok, nrow(rv$completeness_result)))
    
    sc <- rv$completeness_result$server_compare
    n_diff <- sum(grepl("DIFFERENT", sc), na.rm = TRUE)
    n_missing <- sum(grepl("SERVER_MISSING", sc), na.rm = TRUE)
    n_compared <- sum(!is.na(sc) & sc != "SERVER_PATH_NOT_FOUND")
    if (n_compared > 0) {
      add_log(sprintf("Server comparison: %d file(s) differ, %d not found on server, out of %d compared.",
                      n_diff, n_missing, n_compared))
    }
  })
  
  # Independent of run_results/run_btn -- works directly off whatever ARDS
  # files already exist on disk, for the LOA-required program pool. No need
  # to re-run programs just to re-check against the server.
  observeEvent(input$compare_ards_btn, {
    d <- loa_data()
    if (is.null(d) || length(d$all_yy) == 0) {
      showNotification("Load the LOA first (see step 1) so this knows which programs to compare.", type = "warning")
      return(invisible(NULL))
    }
    
    # Strictly use "1. Submission Program Folder"'s own output -- NOT this
    # tool's own local test run. The test run executes under whatever R
    # engine is running this app (e.g. the Linux Posit Workbench engine
    # when accessed via the browser, even through a Windows/Citrix client),
    # which can give numerically different results than a genuine Windows
    # RStudio run (different BLAS/LAPACK, package versions, RNG behavior in
    # multiple-imputation steps, etc.). Comparing the program folder's own
    # pre-existing output against the server avoids that confound.
    source_output_dir <- file.path(normalize_prefix(input$copy_from_path), OUTPUT_SUBDIR)
    server_dir <- input$server_output_path
    server_ok <- !is.null(server_dir) && nzchar(trimws(server_dir)) && dir.exists(server_dir)
    
    # Compute the full itemized detail ONCE per program (not twice) -- the
    # on-screen one-line-per-program summary is then just the first row of
    # that same detail, so both views come from a single comparison pass.
    per_program <- lapply(d$all_yy, function(prog_name) {
      base <- tools::file_path_sans_ext(basename(prog_name))
      ards_name <- paste0(base, "_ards.csv")
      local_csv <- file.path(source_output_dir, ards_name)
      
      detail <- if (!server_ok) {
        data.frame(row = NA_integer_, column = NA_character_,
                   local_value = NA_character_, server_value = NA_character_,
                   status = "SERVER_PATH_NOT_FOUND", stringsAsFactors = FALSE)
      } else {
        compare_ards_to_server_detail(local_csv, file.path(server_dir, ards_name))
      }
      
      summary_result <- if (nrow(detail) == 1 && is.na(detail$row[1])) {
        detail$status[1]
      } else {
        sprintf("DIFFERENT (row %d, col '%s': local='%s' vs server='%s')",
                detail$row[1], detail$column[1], detail$local_value[1], detail$server_value[1])
      }
      
      list(summary = data.frame(program = base, result = summary_result, stringsAsFactors = FALSE),
           detail = cbind(program = base, detail, stringsAsFactors = FALSE))
    })
    
    rv$ards_compare_result <- do.call(rbind, lapply(per_program, `[[`, "summary"))
    rv$ards_compare_detail <- do.call(rbind, lapply(per_program, `[[`, "detail"))
    
    n_match <- sum(rv$ards_compare_result$result == "MATCH")
    add_log(sprintf("Server comparison: %d/%d program(s) matched the server version.",
                    n_match, nrow(rv$ards_compare_result)))
  })
  
  output$ards_compare_table <- renderDT({
    req(rv$ards_compare_result)
    datatable(rv$ards_compare_result, options = list(pageLength = 10), rownames = FALSE)
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
  
  output$save_ards_report_btn <- downloadHandler(
    filename = function() paste0("output_comparison_report_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".csv"),
    content = function(file) {
      req(rv$ards_compare_detail)
      write.csv(rv$ards_compare_detail, file, row.names = FALSE)
    }
  )
}

shinyApp(ui, server)