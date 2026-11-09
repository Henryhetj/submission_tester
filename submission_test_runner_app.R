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
# With no LOA file the program list is everything under programs/ (subfolders such as posthoc included),
# except these folders, which hold helper code / packages rather than programs to run.
NON_PROGRAM_DIRS <- c("renv", "packages", "R_re-usable_functions", "R_re-usable_code_modules")
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
  # Keep the LOA (layout) row order -- not the alphabetical file-listing order --
  # so the table, and therefore the run order, follows the LOA.
  matched_idx <- which(file_base %in% loa_base)
  matched_idx <- matched_idx[order(match(file_base[matched_idx], loa_base))]
  matched <- file_paths[matched_idx]
  unmatched_loa <- loa_names[!(loa_base %in% file_base)]
  list(matched = matched, unmatched_loa_names = unmatched_loa)
}

# Auto-detect the study/protocol folder name from the real "Copy FROM" path
# -- the segment right after the compound code (matches "ly" + digits, e.g.
# "ly3437943", generically -- not hardcoded to one specific compound so this
# works for other studies/compounds too) and before the next "/". Falls back
# to DEFAULT_PROTOCOL_ID if that pattern isn't found in the path.
# The "1. Submission Program Folder" box may point at the final folder itself
# or at final/programs (or anything below final): find the LAST "/final"
# segment in the path and use everything up to it. The last one, because a
# path can contain more than one (e.g. .../final/programs/.../final/programs).
# The whole final folder is what gets copied -- the programs read data/ and
# documentation/ -- and output/, logs/ are looked up under it too.
resolve_final_path <- function(p) {
  p <- normalize_prefix(gsub("\\\\", "/", p))
  m <- gregexpr("(?i)/final(?=/|$)", p, perl = TRUE)[[1]]
  if (m[1] == -1L) return(p)            # no "final" segment: use the path as typed
  last <- length(m)
  substr(p, 1L, m[last] + attr(m, "match.length")[last] - 1L)
}

list_program_files <- function(prog_dir, skip_dirs = character(0)) {
  if (!dir.exists(prog_dir)) return(character(0))
  pat <- "\\.R$"
  top <- list.files(prog_dir, pattern = pat, ignore.case = TRUE, full.names = TRUE)
  subs <- list.dirs(prog_dir, recursive = FALSE, full.names = TRUE)
  subs <- subs[!(tolower(basename(subs)) %in% tolower(skip_dirs))]
  deeper <- unlist(lapply(subs, function(d)
    list.files(d, pattern = pat, ignore.case = TRUE, recursive = TRUE, full.names = TRUE)))
  c(top, deeper)
}

resolve_subdir_ci <- function(base, rel) {
  parts <- strsplit(gsub("\\\\", "/", trimws(rel)), "/")[[1]]
  parts <- parts[nzchar(parts)]
  cur <- base
  for (p in parts) {
    subs <- if (dir.exists(cur)) list.dirs(cur, recursive = FALSE, full.names = FALSE) else character(0)
    hit <- subs[tolower(subs) == tolower(p)]
    cur <- file.path(cur, if (length(hit) > 0) hit[1] else p)
  }
  cur
}

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

# Find a file by name in a folder, or else in one of its immediate subfolders
# (e.g. output/posthoc). Returns the path, or NA if there is none.
find_file_shallow <- function(dir, fname) {
  if (is.null(dir) || !nzchar(dir) || !dir.exists(dir)) return(NA_character_)
  flat <- file.path(dir, fname)
  if (file.exists(flat)) return(flat)
  subs <- list.dirs(dir, recursive = FALSE, full.names = TRUE)
  hit <- file.path(subs, fname)
  hit <- hit[file.exists(hit)]
  if (length(hit) > 0) hit[1] else NA_character_
}

# Guess the expected output file(s) for a program using a naming heuristic:
# the folder itself first, then its immediate subfolders (e.g. output/posthoc).
guess_expected_outputs <- function(program_file, output_dir, output_extensions) {
  base <- tools::file_path_sans_ext(basename(program_file))
  vapply(paste0(base, ".", output_extensions), function(f) {
    p <- find_file_shallow(output_dir, f)
    if (is.na(p)) file.path(output_dir, f) else p
  }, character(1), USE.NAMES = FALSE)
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
  
  # Look at BOTH sides before calling anything "missing": a program that
  # produces no ARDS file anywhere matches the server (neither has one).
  local_exists <- file.exists(local_csv)
  server_exists <- file.exists(server_csv)
  if (!local_exists && !server_exists) return(no_diff_row("MATCH (no ARDS file on either side)"))
  if (!local_exists) return(no_diff_row("LOCAL_MISSING"))
  if (!server_exists) return(no_diff_row("SERVER_MISSING"))
  
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
    # column's differences too, not just RESULT's. The list below stays in
    # that order (RESULT first, then the other columns in file order), and
    # within each column the rows are ascending.
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
    rownames(out) <- NULL
    out
  }, error = function(e) {
    no_diff_row(paste0("COMPARE_FAILED (", conditionMessage(e), ")"))
  })
}

# Thin wrapper around compare_ards_to_server_detail() for the on-screen
# summary table -- one line per program, showing just the first difference
# (a RESULT difference whenever there is one). The full itemized list is what the
# downloadable report uses instead of this summary.
compare_ards_to_server <- function(local_csv, server_csv,
                                   ignore_cols = c("PROGRAM", "OUTPUT"), tol = 1e-6) {
  d <- compare_ards_to_server_detail(local_csv, server_csv, ignore_cols, tol)
  if (nrow(d) == 1 && is.na(d$row[1])) return(d$status[1])
  i <- 1
  sprintf("DIFFERENT (row %d, col '%s': local='%s' vs server='%s')",
          d$row[i], d$column[i], d$local_value[i], d$server_value[i])
}

# ---------------------------------------------------------------------------
# Comparison report as an editable .xlsx workbook
#   Summary     : one row per program (+ an ALL PROGRAMS row on top) -- MATCH or not
#   Differences : every differing cell (RESULT first, then other columns);
#                 a single "No differences found" line when everything matches
# ---------------------------------------------------------------------------
build_ards_report_sheets <- function(detail) {
  progs <- unique(detail$program)
  per <- lapply(progs, function(p) {
    d <- detail[detail$program == p, , drop = FALSE]
    has_cells <- any(!is.na(d$row))
    data.frame(program = p,
               result = if (has_cells) "DIFFERENT" else d$status[1],
               differing_cells = sum(!is.na(d$row)),
               stringsAsFactors = FALSE)
  })
  summary <- do.call(rbind, per)
  not_match <- !startsWith(summary$result, "MATCH")
  overall <- data.frame(
    program = "ALL PROGRAMS",
    result = if (any(not_match)) sprintf("DIFFERENT (%d of %d programs not matching)", sum(not_match), nrow(summary)) else "MATCH",
    differing_cells = sum(summary$differing_cells),
    stringsAsFactors = FALSE)
  diffs <- detail[!startsWith(detail$status, "MATCH"), , drop = FALSE]
  if (nrow(diffs) == 0) {
    diffs <- data.frame(message = "No differences found (all programs match).", stringsAsFactors = FALSE)
  }
  list(Summary = rbind(overall, summary), Differences = diffs)
}

# Which xlsx writer is installed on this server (NA if none)
xlsx_engine <- function() {
  if (requireNamespace("openxlsx", quietly = TRUE)) "openxlsx"
  else if (requireNamespace("writexl", quietly = TRUE)) "writexl"
  else NA_character_
}

write_report_xlsx <- function(sheets, file) {
  if (identical(xlsx_engine(), "openxlsx")) {
    wb <- openxlsx::createWorkbook()
    hdr <- openxlsx::createStyle(textDecoration = "bold", fgFill = "#D9E1F2")
    for (nm in names(sheets)) {
      df <- sheets[[nm]]
      openxlsx::addWorksheet(wb, nm)
      openxlsx::writeData(wb, nm, df, headerStyle = hdr)
      openxlsx::freezePane(wb, nm, firstRow = TRUE)
      openxlsx::addFilter(wb, nm, rows = 1, cols = seq_len(ncol(df)))
      openxlsx::setColWidths(wb, nm, cols = seq_len(ncol(df)), widths = "auto")
    }
    openxlsx::saveWorkbook(wb, file, overwrite = TRUE)
  } else {
    writexl::write_xlsx(sheets, file)
  }
  invisible(file)
}

log_line <- function(msg) {
  paste0("[", format(Sys.time(), "%Y-%m-%d %H:%M:%S"), "] ", msg)
}

# Start the folder copy as an async process, OS-appropriate:
#   Windows -> robocopy (mirrors subfolders, has a real /LOG file)
#   Linux   -> `cp -rv`  (robocopy doesn't exist on Linux/Workbench)
# Returns a processx::process object in both cases so the calling code
# (the copy_timer poller) doesn't need to know which OS it's on.
start_copy_process <- function(source, dest, log_path, update_only = TRUE) {
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
    # -u: only copy files that are missing in dest or newer in the source, so a
    # re-run just picks up new/changed files (e.g. a program added to programs/posthoc)
    args = c(if (update_only) "-ruv" else "-rv", src_contents, dest),
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
bootstrap_packages <- function(programs_dir, lib, repos, scan_files, use_pins = TRUE,
                               pins_file = "install_packages.R", pin_only = "", pin_compiled = FALSE) {
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
  
  # Match the versions recorded in the submission's own version list (default
  # install_packages.R, in the programs folder -- the file name is configurable),
  # for packages the programs use. Only pure-R packages (NeedsCompilation: no)
  # are swapped unless pin_compiled is TRUE (compiled ones then need a compiler
  # toolchain such as Rtools). pin_only limits this to the named packages.
  pins_path <- file.path(programs_dir, pins_file)
  if (isTRUE(use_pins)) {
    if (!file.exists(pins_path)) {
      cat(sprintf("[bootstrap] version list '%s' not found in the programs folder -- versions are not matched\n", pins_file))
    } else {
      sl <- readLines(pins_path, warn = FALSE)
      mm <- regmatches(sl, regexec('^\\s*"?([A-Za-z][A-Za-z0-9.]*)"?\\s*=\\s*"([0-9][^"]*)"\\s*,?\\s*(#.*)?$', sl))
      mm <- mm[lengths(mm) == 4]
      pins <- setNames(vapply(mm, function(z) z[3], ""), vapply(mm, function(z) z[2], ""))
      snap <- NA_character_
      cl <- grep("CRAN_URL\\s*<-", sl, value = TRUE)
      if (length(cl) > 0 && grepl("/cran/[0-9-]+", cl[1])) {
        snap <- sub('^.*/cran/([0-9-]+).*$', "\\1", cl[1])
      }
      only <- trimws(strsplit(pin_only, "[,; ]+")[[1]])
      only <- only[nzchar(only)]
      if (length(only) > 0) {
        unknown <- setdiff(only, names(pins))
        if (length(unknown) > 0) cat("[bootstrap] not in the version list, ignored:", paste(unknown, collapse = ", "), "\n")
        targets <- setdiff(intersect(only, names(pins)), c(builtin, local_names))
      } else {
        targets <- setdiff(intersect(names(pins), c(used, local_deps)), c(builtin, local_names))
      }
      cat(sprintf("[bootstrap] %s pins %d package version(s); checking %d package(s)%s\n",
                  pins_file, length(pins), length(targets), if (length(only) > 0) " (limited to the ones you listed)" else " used by these programs"))
      base1 <- unname(repos[1])
      bases <- unique(c(base1,
                        if (!is.na(snap)) sub("/latest$", paste0("/", snap), base1),
                        sub("__linux__/[^/]+/", "", base1),
                        if (!is.na(snap)) sub("__linux__/[^/]+/", "", sub("/latest$", paste0("/", snap), base1))))
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
        if (isTRUE(tolower(nc) == "yes") && !isTRUE(pin_compiled)) {
          cat(sprintf("[bootstrap]   %s: pinned %s needs compiling, kept %s\n", p, ver, cur))
          next
        }
        tryCatch({
          install.packages(tarball, lib = lib, repos = NULL, type = "source")
          now <- tryCatch(as.character(packageVersion(p, lib.loc = lib)), error = function(e) NA_character_)
          if (!is.na(now) && package_version(now) == package_version(ver)) {
            cat(sprintf("[bootstrap]   %s: %s -> pinned %s\n", p, cur, ver))
          } else {
            cat(sprintf("[bootstrap]   %s: pin to %s did not take effect (still %s)\n", p, ver, now))
          }
        }, error = function(e) cat(sprintf("[bootstrap]   %s: pin to %s failed (%s)\n", p, ver, conditionMessage(e))))
      }
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
bootstrap_script_lines <- function(programs_dir, lib_expr, repos, scan_files, use_pins = TRUE,
                                   pins_file = "install_packages.R", pin_only = "", pin_compiled = FALSE) {
  c(paste0("bootstrap_packages <- ", paste(deparse(bootstrap_packages, width.cutoff = 500L), collapse = "\n")),
    sprintf("bootstrap_packages(%s, %s, %s, %s, %s, %s, %s, %s)", deparse(programs_dir), lib_expr,
            paste(deparse(repos), collapse = ""), paste(deparse(scan_files), collapse = ""),
            deparse(isTRUE(use_pins)), deparse(as.character(pins_file)[1]),
            deparse(paste(as.character(pin_only), collapse = ",")), deparse(isTRUE(pin_compiled))))
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
  tags$head(tags$script(HTML(paste0(
    "$(document).on('dblclick', '#br_list option', function() {",
    "  Shiny.setInputValue('br_dbl', $(this).val(), {priority: 'event'});",
    "});")))),
  
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
    div(style = "display:flex; gap:6px; align-items:flex-start;",
        div(style = "flex:1; min-width:0;",
            textInput("copy_from_path", NULL, width = "100%",
                      value = if (IS_WINDOWS) {
                        "Z:/qa/ly3437943/j1i_mc_gzbk/common/documentation/submission/submission_files/j1i_mc_gzbk/final"
                      } else {
                        "/lillyce/qa/ly3437943/j1i_mc_gzbk/common/documentation/submission/submission_files/j1i_mc_gzbk/final"
                      })),
        actionButton("browse_copy_from", "Browse", class = "btn-outline-secondary btn-sm", style = "width:auto; margin-top:2px;")),
    helpText("Path to the final folder or final/programs (read-only)."),
    
    h6("2. LOA file (optional)"),
    div(style = "display:flex; gap:6px; align-items:flex-start;",
        div(style = "flex:1; min-width:0;",
            textInput("loa_path", NULL, width = "100%",
                      value = "/lillyce/qa/ly3437943/j1i_mc_gzbk/final/documentation/loa_trackers/gzbk_final_TFL_link.xlsx")),
        actionButton("browse_loa", "Browse", class = "btn-outline-secondary btn-sm", style = "width:auto; margin-top:2px;")),
    helpText("Limits the run to Y/Y programs. Empty = list all programs under programs/."),
    
    h6("3. Study/protocol folder (auto-detected)"),
    verbatimTextOutput("hardcoded_path_breakdown"),
    helpText("Taken from the path above."),
    
    h6("4. Server output folder"),
    div(style = "display:flex; gap:6px; align-items:flex-start;",
        div(style = "flex:1; min-width:0;",
            textInput("server_output_path", NULL, width = "100%",
                      value = "/lillyce/prd/ly3437943/j1i_mc_gzbk/final/output/shared")),
        actionButton("browse_server", "Browse", class = "btn-outline-secondary btn-sm", style = "width:auto; margin-top:2px;")),
    helpText("Server outputs to compare against."),
    bslib::accordion(
      open = FALSE,
      bslib::accordion_panel(
        "Advanced settings",
        h6("R engine"),
        radioButtons("r_engine_mode", NULL,
                     choices = c("Server R engine (Linux, default)" = "default",
                                 "Local R engine (Windows)" = "custom"),
                     selected = "custom"),
        conditionalPanel(
          condition = "input.r_engine_mode == 'custom'",
          div(style = paste0("background:#fff3cd; border:1px solid #ffe69c; ",
                             "border-radius:6px; padding:8px 12px; margin-bottom:8px; font-size:0.8rem; color:#664d03;"),
              tags$strong("Local R must already be installed. "),
              "'Run selected programs' creates run_selected_programs.bat -- double-click it to start the run. ",
              "After switching engines, run 'Set up testing environment' again (with 'Force re-copy')."),
          textInput("rscript_path", "Rscript path (not RStudio.exe)", value = "Y:/Programs/R/R-4.5.2/bin/Rscript.exe"),
          helpText("Packages install into <drive>:\\", WIN_LIB_DIRNAME, "_<R version> (drive from the Rscript path).")
        ),
        checkboxInput("clean_library", "Clean package library (Linux engine)", value = TRUE),
        helpText("Ignores pre-installed packages; each run starts clean."),
        checkboxInput("use_pinned_versions", "Use pinned package versions (slower)", value = TRUE),
        conditionalPanel(
          condition = "input.use_pinned_versions",
          textInput("pins_file", "Version list file (in the programs folder)", value = "install_packages.R"),
          textInput("pin_only", "Only these packages (comma-separated; empty = all used)", value = "officer, flextable"),
          checkboxInput("pin_compiled", "Also compiled packages (needs Rtools; slow)", value = FALSE),
          helpText("Versions are read from the file, never run.")
        ),
        actionButton("reset_clean_lib", "Reset clean library", style = "width: auto;"),
        hr(),
        h6("LOA column names"),
        textInput("loa_sheet_name", "Sheet name (Excel files only)", value = "TFL Metadata"),
        textInput("loa_status_col", "Status column name", value = "Program for submission/ Executable?"),
        textInput("loa_program_col", "Program-name column name", value = "Program Name"),
        helpText("Matching ignores case/spacing; falls back to the first sheet."),
        textInput("posthoc_subdir", "Only this folder inside programs/ when there is no LOA (empty = all of programs/)", value = ""),
        hr(),
        h6("Path replacement"),
        textInput("file_extensions", "File extensions to scan/replace", value = "R"),
        checkboxInput("ignore_case", "Ignore case when matching path", value = FALSE),
        checkboxInput("keep_backup", "Keep an untouched backup before replacing", value = FALSE),
        checkboxInput("force_recopy", "Re-copy everything, not just new or changed files", value = FALSE)
      )
    ),
    actionButton("confirm_btn", "Confirm settings", class = "btn-primary btn-sm", style = "width: auto; margin-top: 8px;"),
    uiOutput("cfg_status")
  ),
  
  bslib::layout_columns(
    col_widths = c(6, 6),
    
    bslib::card(
      full_screen = TRUE,
      height = "780px",
      bslib::card_header("1. Set Up Testing Environment & Run Programs"),
      div(style = "display: flex; flex-direction: row; align-items: center; flex-wrap: wrap; gap: 8px 16px;",
          actionButton("setup_env_btn", "Set up testing environment", class = "btn-primary btn-sm", style = "width: auto;"),
          actionButton("refresh_programs_btn", "Refresh program list", class = "btn-primary btn-sm", style = "width: auto;"),
          div(style = "white-space: nowrap;", checkboxInput("recurse_subfolders", "Include subfolders", value = FALSE))
      ),
      uiOutput("programs_help"),
      h6(textOutput("programs_select_heading", inline = TRUE)),
      div(style = "display: flex; flex-direction: row; gap: 8px; margin-bottom: 4px;",
          actionButton("select_all_btn", "Select all", class = "btn-outline-primary btn-sm", style = "width: auto;"),
          actionButton("deselect_all_btn", "Deselect all", class = "btn-outline-secondary btn-sm", style = "width: auto;"),
          actionButton("confirm_selection_btn", "Confirm selection", class = "btn-primary btn-sm", style = "width: auto;")
      ),
      uiOutput("selection_status"),
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
          downloadButton("save_ards_report_btn", "Download report (xlsx)", class = "btn-sm", style = "width: auto;")
      ),
      helpText("Compares the confirmed programs' output with the server."),
      DTOutput("ards_compare_table"),
      
      hr(),
      actionButton("clear_cache_btn", "Clear testing environment", class = "btn-danger btn-sm", style = "width: auto;"),
      helpText("Removes the local test copy."),
      
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
    cfg = NULL,                # confirmed Setup settings (set by the Confirm button)
    selected = NULL,           # confirmed program selection (full paths, in run order)
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
  # Everything downstream uses the settings captured when "Confirm settings" was
  # clicked (rv$cfg), not the live text boxes, so a half-edited path can't change
  # what the buttons below act on.
  adv_inputs <- c("r_engine_mode", "rscript_path", "clean_library", "use_pinned_versions", "pins_file", "pin_only",
                  "pin_compiled", "loa_sheet_name", "loa_status_col", "loa_program_col", "posthoc_subdir",
                  "file_extensions", "ignore_case", "keep_backup", "force_recopy")
  live_settings <- function() {
    adv <- lapply(adv_inputs, function(n) { v <- input[[n]]; if (is.character(v)) trimws(v) else v })
    names(adv) <- adv_inputs
    if (is.null(adv$posthoc_subdir)) adv$posthoc_subdir <- ""
    c(list(copy_from_path = trimws(input$copy_from_path),
           loa_path = trimws(input$loa_path),
           server_output_path = trimws(input$server_output_path)),
      adv)
  }
  
  need_cfg <- function() {
    if (is.null(rv$cfg)) {
      showNotification("Click 'Confirm settings' in Setup first.", type = "warning")
      return(FALSE)
    }
    TRUE
  }
  
  # The programs currently highlighted in the table, always in table order
  # (= run order), whatever order the rows were clicked in.
  live_selection <- function() {
    sel <- input$programs_table_rows_selected
    if (length(sel) == 0) character(0) else rv$programs[sort(sel)]
  }
  
  # Run and the Output & Log checks all act on the CONFIRMED selection.
  need_selection <- function() {
    if (is.null(rv$selected) || length(rv$selected) == 0) {
      showNotification("Select the programs in the table and click 'Confirm selection' first.", type = "warning")
      return(FALSE)
    }
    if (!identical(live_selection(), rv$selected)) {
      showNotification("The selection changed since it was confirmed. Click 'Confirm selection' again.", type = "warning")
      return(FALSE)
    }
    TRUE
  }
  
  # ---- Browse buttons: a small server-side folder/file picker -------------------
  br <- reactiveValues(target = NULL, mode = "dir", dir = NULL)
  
  browse_entries <- function(dir, mode) {
    dirs <- tryCatch(list.dirs(dir, recursive = FALSE, full.names = FALSE), error = function(e) character(0))
    dirs <- sort(dirs[!startsWith(dirs, ".")])
    items <- c("-- this folder --" = "")
    if (length(dirs) > 0) items <- c(items, setNames(file.path(dir, dirs), paste0("[folder] ", dirs)))
    if (identical(mode, "file")) {
      fl <- tryCatch(list.files(dir, full.names = FALSE), error = function(e) character(0))
      fl <- fl[grepl("\\.(xlsx|xls|csv)$", fl, ignore.case = TRUE) & !dir.exists(file.path(dir, fl))]
      if (length(fl) > 0) items <- c(items, setNames(file.path(dir, fl), fl))
    }
    items
  }
  
  open_browser <- function(target, mode) {
    cur <- trimws(input[[target]])
    start <- if (identical(mode, "file") && nzchar(cur)) dirname(cur) else cur
    while (nzchar(start) && !dir.exists(start) && dirname(start) != start) start <- dirname(start)
    if (!nzchar(start) || !dir.exists(start)) start <- normalize_prefix(DEFAULT_DEST_ROOT)
    br$target <- target; br$mode <- mode; br$dir <- start
    showModal(modalDialog(
      title = if (identical(mode, "dir")) "Select a folder" else "Select a file",
      textInput("br_path", NULL, value = start, width = "100%"),
      div(style = "display:flex; gap:6px; margin-bottom:8px;",
          actionButton("br_up", "Up one level", class = "btn-sm"),
          actionButton("br_go", "Go to this path", class = "btn-sm")),
      selectInput("br_list", NULL, choices = browse_entries(start, mode), selected = "",
                  size = 12, selectize = FALSE, width = "100%"),
      helpText(if (identical(mode, "dir")) "Double-click or 'Open folder' to go in. 'Select' uses the highlighted folder, or this folder if none is highlighted."
               else "Double-click or 'Open folder' to go into a folder. Highlight a file and click 'Select'."),
      footer = tagList(actionButton("br_open", "Open folder", class = "btn-sm"),
                       actionButton("br_select", "Select", class = "btn-primary btn-sm"),
                       modalButton("Cancel")),
      size = "l", easyClose = TRUE))
  }
  
  br_goto <- function(path) {
    if (!nzchar(path) || !dir.exists(path)) {
      showNotification("Folder not found.", type = "warning")
      return(invisible(NULL))
    }
    br$dir <- path
    updateTextInput(session, "br_path", value = path)
    updateSelectInput(session, "br_list", choices = browse_entries(path, br$mode), selected = "")
  }
  
  br_finish <- function(path) {
    updateTextInput(session, br$target, value = path)
    removeModal()
  }
  
  observeEvent(input$browse_copy_from, open_browser("copy_from_path", "dir"))
  observeEvent(input$browse_loa, open_browser("loa_path", "file"))
  observeEvent(input$browse_server, open_browser("server_output_path", "dir"))
  observeEvent(input$br_up, br_goto(dirname(br$dir)))
  observeEvent(input$br_go, br_goto(trimws(input$br_path)))
  observeEvent(input$br_open, {
    sel <- input$br_list
    if (length(sel) == 1 && nzchar(sel) && dir.exists(sel)) br_goto(sel)
  })
  observeEvent(input$br_dbl, {
    sel <- input$br_dbl
    if (length(sel) == 1 && nzchar(sel)) {
      if (dir.exists(sel)) br_goto(sel) else if (identical(br$mode, "file")) br_finish(sel)
    }
  })
  observeEvent(input$br_select, {
    sel <- input$br_list
    has_sel <- length(sel) == 1 && !is.na(sel) && nzchar(sel)
    if (identical(br$mode, "dir")) {
      br_finish(if (has_sel && dir.exists(sel)) sel else br$dir)
    } else if (has_sel && !dir.exists(sel)) {
      br_finish(sel)
    } else {
      showNotification("Pick a file from the list (folders open with 'Open folder').", type = "warning")
    }
  })
  
  protocol_id_reactive <- reactive({
    req(rv$cfg)
    extracted <- extract_protocol_id_from_path(rv$cfg$copy_from_path)
    if (is.na(extracted)) DEFAULT_PROTOCOL_ID else extracted
  })
  
  source_final_path <- reactive({
    req(rv$cfg)
    resolve_final_path(rv$cfg$copy_from_path)
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
    if (identical(rv$cfg$r_engine_mode, "custom")) rv$cfg$rscript_path else DEFAULT_RSCRIPT_PATH
  })
  
  observeEvent(input$clear_cache_btn, {
    if (!need_cfg()) return(invisible(NULL))
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
    pth <- input$copy_from_path
    fp <- resolve_final_path(pth)
    id <- extract_protocol_id_from_path(pth)
    if (is.na(id)) id <- DEFAULT_PROTOCOL_ID
    sprintf("Detected: %s\nFinal folder: %s%s", id, fp,
            if (!grepl("(?i)/final$", fp, perl = TRUE)) "\n(no 'final' folder found in the path -- used as typed)" else "")
  })
  
  observeEvent(input$confirm_btn, {
    new_cfg <- live_settings()
    if (!nzchar(new_cfg$copy_from_path)) {
      showNotification("Enter the submission program folder (1) first.", type = "error")
      return(invisible(NULL))
    }
    id_of <- function(p) {
      e <- extract_protocol_id_from_path(p)
      if (is.na(e)) DEFAULT_PROTOCOL_ID else e
    }
    old_id <- if (is.null(rv$cfg)) NA_character_ else id_of(rv$cfg$copy_from_path)
    new_id <- id_of(new_cfg$copy_from_path)
    rv$cfg <- new_cfg
    if (!is.na(old_id) && !identical(old_id, new_id)) {
      rv$programs <- character(0)
      rv$selected <- NULL
      loa_data(NULL)
      add_log("Study changed -- the program list was cleared. Click 'Refresh program list'.")
    }
    
    final_src <- resolve_final_path(new_cfg$copy_from_path)
    dest <- file.path(normalize_prefix(DEFAULT_DEST_ROOT),
                      extract_relative_structure(paste0("C:/submission_files/", new_id, "/final")))
    add_log("Settings confirmed:")
    add_log(sprintf("  Submission folder: %s%s", final_src, if (dir.exists(final_src)) "" else "  (not reachable from here)"))
    add_log(sprintf("  Working folder: %s%s", dest, if (dir.exists(dest)) "  (exists)" else "  (not set up yet: click 'Set up testing environment')"))
    add_log(sprintf("  Program list: %s", if (nzchar(new_cfg$loa_path)) {
      paste0("from the LOA ", new_cfg$loa_path, if (file.exists(new_cfg$loa_path)) "" else "  (file not found)")
    } else {
      if (nzchar(new_cfg$posthoc_subdir)) sprintf("no LOA -- programs in %s/%s of the working folder", PROGRAMS_SUBDIR, new_cfg$posthoc_subdir)
      else sprintf("no LOA -- all programs under %s/ of the working folder (subfolders included)", PROGRAMS_SUBDIR)
    }))
    add_log(sprintf("  Server output folder: %s%s", new_cfg$server_output_path,
                    if (dir.exists(new_cfg$server_output_path)) "" else "  (not found)"))
    add_log(sprintf("  R engine: %s", if (identical(new_cfg$r_engine_mode, "custom"))
      paste0("local Windows R (", new_cfg$rscript_path, ")") else "server R engine (Linux)"))
    showNotification("Settings confirmed.", type = "message")
  })
  
  output$cfg_status <- renderUI({
    box_style <- function(bg, bd, fg) paste0("background:", bg, "; border:1px solid ", bd, "; border-radius:6px; ",
                                             "padding:8px 10px; margin:8px 0; font-size:0.8rem; color:", fg, ";")
    if (is.null(rv$cfg)) {
      return(div(style = box_style("#fff3cd", "#ffe69c", "#664d03"), "Not confirmed yet. Click 'Confirm settings'."))
    }
    if (!identical(live_settings(), rv$cfg)) {
      return(div(style = box_style("#fff3cd", "#ffe69c", "#664d03"),
                 "Settings changed since the last confirmation. Click 'Confirm settings' again."))
    }
    div(style = box_style("#d1e7dd", "#badbcc", "#0f5132"),
        tags$strong("Confirmed"),
        tags$br(), paste0("Submission folder: ", resolve_final_path(rv$cfg$copy_from_path)),
        tags$br(), paste0("R engine: ", if (identical(rv$cfg$r_engine_mode, "custom")) "local Windows R" else "server R (Linux)"),
        tags$br(), paste0("Program list: ", if (nzchar(rv$cfg$loa_path)) "from the LOA" else
          if (nzchar(rv$cfg$posthoc_subdir)) paste0("no LOA, ", PROGRAMS_SUBDIR, "/", rv$cfg$posthoc_subdir)
          else paste0("no LOA, all of ", PROGRAMS_SUBDIR, "/")))
  })
  
  # -------------------------------------------------------------------------
  # Tab 1: Copy
  # -------------------------------------------------------------------------
  
  run_copy_step <- function() {
    src <- source_final_path()
    dst <- effective_dest_root()
    
    if (!dir.exists(src)) {
      add_log(paste0("ERROR: source path does not exist or is not reachable: ", src))
      showNotification("Source path not found. Check the path and network connectivity.", type = "error")
      return(FALSE)
    }
    
    rv$copy_total_files <- length(list.files(src, recursive = TRUE))
    add_log(sprintf("Setting up files from %s (%d found; %s)...", src, rv$copy_total_files,
                    if (isTRUE(rv$cfg$force_recopy)) "copying everything" else "copying only new or changed files"))
    
    log_path <- file.path(tempdir(), paste0("copy_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".log"))
    rv$copy_log_path <- log_path
    
    proc <- tryCatch({
      start_copy_process(src, dst, log_path, update_only = !isTRUE(rv$cfg$force_recopy))
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
    exts <- trimws(strsplit(rv$cfg$file_extensions, ",")[[1]])
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
    scan_res <- scan_for_drive_refs(files, pat, ignore_case = rv$cfg$ignore_case)
    scan_res$file <- strip_prefix_literal(scan_res$file, root)
    rv$scan_result <- scan_res[order(-scan_res$total), ]
    add_log(sprintf("Scan complete. %d file(s) contain at least one match.", sum(scan_res$total > 0, na.rm = TRUE)))
    
    if (identical(rv$cfg$r_engine_mode, "custom")) {
      rv$replace_result <- NULL
      add_log("Windows engine selected: programs are left exactly as delivered (no path replacement).")
      if (sum(scan_res$total, na.rm = TRUE) == 0) {
        add_log(paste0("WARNING: no file still contains '", hardcoded_path(), "' -- this copy looks like it was already modified ",
                       "(for example by an earlier set-up with the Linux engine). Tick 'Force re-copy' in Advanced settings and ",
                       "run 'Set up testing environment' again to restore the originals."))
      }
      return(invisible(NULL))
    }
    
    # ...then apply the replacement, appending n_replaced/status onto it.
    if (isTRUE(rv$cfg$keep_backup)) {
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
    apply_res <- apply_drive_replace(files, pat, DEFAULT_DEST_ROOT, ignore_case = rv$cfg$ignore_case)
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
  
  loa_given <- function() isTRUE(nzchar(rv$cfg$loa_path))
  
  no_loa_subdir <- function() {
    if (is.null(rv$cfg$posthoc_subdir)) "" else rv$cfg$posthoc_subdir
  }
  
  refresh_programs <- function() {
    prog_dir <- file.path(effective_dest_root(), PROGRAMS_SUBDIR)
    sub <- no_loa_subdir()
    if (loa_given()) {
      files <- list_files_by_ext(prog_dir, c("R"), recursive = isTRUE(input$recurse_subfolders))
      where <- if (isTRUE(input$recurse_subfolders)) " (including subfolders)" else ""
    } else {
      # No LOA: everything under programs/ (subfolders such as posthoc included),
      # or just the one folder named in Advanced settings.
      scope_dir <- if (nzchar(sub)) resolve_subdir_ci(prog_dir, sub) else prog_dir
      files <- list_program_files(scope_dir, NON_PROGRAM_DIRS)
      where <- sprintf(" under %s (no LOA given)", scope_dir)
    }
    rv$programs <- files
    rv$selected <- NULL
    loa_data(NULL)  # a fresh file listing invalidates any prior LOA-derived order/selection
    add_log(sprintf("Found %d program(s)%s", length(files), where))
    
    if (!loa_given()) {
      # Programs that exist in the SOURCE folder but are not in the working copy yet
      src_prog <- file.path(source_final_path(), PROGRAMS_SUBDIR)
      if (dir.exists(src_prog)) {
        rel <- function(f, root) substring(f, nchar(root) + 2L)
        miss <- setdiff(rel(list_program_files(src_prog, NON_PROGRAM_DIRS), src_prog),
                        rel(list_program_files(prog_dir, NON_PROGRAM_DIRS), prog_dir))
        if (length(miss) > 0) {
          add_log(sprintf("%d program(s) are in the source folder but not in the working copy yet: %s. Click 'Set up testing environment' to copy them.",
                          length(miss), paste(utils::head(miss, 5), collapse = ", ")))
          showNotification(sprintf("%d program(s) exist in the source folder but not in the working copy yet. Click 'Set up testing environment' to copy them.",
                                   length(miss)), type = "warning", duration = 15)
        }
      }
      if (length(files) == 0) {
        subs <- if (dir.exists(prog_dir)) list.dirs(prog_dir, recursive = FALSE, full.names = FALSE) else character(0)
        add_log(sprintf("Looked in: %s -- %s.", if (nzchar(sub)) resolve_subdir_ci(prog_dir, sub) else prog_dir,
                        if (dir.exists(prog_dir)) "no .R programs found" else "the working folder does not exist yet (click 'Set up testing environment')"))
        add_log(sprintf("Subfolders of %s: %s", prog_dir, if (length(subs) > 0) paste(subs, collapse = ", ") else "(none)"))
        showNotification("No programs found in the working copy. The Console Log shows what was checked.", type = "warning", duration = 12)
      }
    }
  }
  
  observeEvent(input$setup_env_btn, {
    if (!need_cfg()) return(invisible(NULL))
    # Copy step, then prepare. The copy is incremental: only files that are
    # missing or newer in the source are copied, so re-running Set up is cheap
    # and picks up newly added programs. Tick "Re-copy everything" in Advanced
    # settings to overwrite every file instead.
    rv$auto_chain_pending <- TRUE
    ok <- run_copy_step()
    if (!ok) rv$auto_chain_pending <- FALSE
  })
  
  observeEvent(input$refresh_programs_btn, {
    if (!need_cfg()) return(invisible(NULL))
    refresh_programs()
    if (loa_given()) load_loa_and_match()
  })
  
  output$programs_help <- renderUI({
    if (is.null(rv$cfg)) return(helpText("Confirm the settings in Setup first."))
    if (loa_given()) {
      helpText("Lists the programs flagged Y/Y in the LOA.")
    } else {
      helpText(if (nzchar(no_loa_subdir())) sprintf("No LOA: lists the programs in %s/%s.", PROGRAMS_SUBDIR, no_loa_subdir())
               else sprintf("No LOA: lists all programs under %s/ (subfolders such as posthoc included).", PROGRAMS_SUBDIR))
    }
  })
  
  output$programs_table <- renderDT({
    d <- loa_data()
    to_be_run <- if (!is.null(d)) rv$programs %in% d$ordered else rep(FALSE, length(rv$programs))
    no_loa <- !is.null(rv$cfg) && !loa_given()
    shown <- if (no_loa) substring(rv$programs, nchar(file.path(effective_dest_root(), PROGRAMS_SUBDIR)) + 2L) else basename(rv$programs)
    df <- data.frame(program = shown, stringsAsFactors = FALSE)
    # Row highlighting IS "will be run" -- no separate "To be run" column
    # that could get out of sync with it. LOA matches just set the initial
    # selection; the person can then freely check/uncheck rows, and
    # whatever's actually highlighted when they click Run is what runs.
    # No LOA: nothing is pre-selected -- the person picks what to run.
    selected_idx <- if (!is.null(d)) which(to_be_run) else if (no_loa) integer(0) else seq_len(nrow(df))
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
    req(rv$cfg$loa_path)
    if (!file.exists(rv$cfg$loa_path)) {
      showNotification("LOA file not found at that path.", type = "error")
      add_log(paste0("LOA load failed: file not found: ", rv$cfg$loa_path))
      return(invisible(NULL))
    }
    
    res <- tryCatch(
      read_loa_yy_programs(rv$cfg$loa_path, rv$cfg$loa_status_col, rv$cfg$loa_program_col, rv$cfg$loa_sheet_name),
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
  
  observeEvent(input$confirm_selection_btn, {
    if (!need_cfg()) return(invisible(NULL))
    sel <- live_selection()
    if (length(sel) == 0) {
      showNotification("Select at least one program first.", type = "warning")
      return(invisible(NULL))
    }
    rv$selected <- sel
    add_log(sprintf("Selection confirmed: %d program(s), in this order: %s%s", length(sel),
                    paste(utils::head(basename(sel), 3), collapse = ", "),
                    if (length(sel) > 3) ", ..." else ""))
  })
  
  output$selection_status <- renderUI({
    box_style <- function(bg, bd, fg) paste0("background:", bg, "; border:1px solid ", bd, "; border-radius:6px; ",
                                             "padding:6px 10px; margin:4px 0 8px 0; font-size:0.8rem; color:", fg, ";")
    if (is.null(rv$selected)) {
      return(div(style = box_style("#fff3cd", "#ffe69c", "#664d03"),
                 "Selection not confirmed. Pick the programs, then click 'Confirm selection'."))
    }
    if (!identical(live_selection(), rv$selected)) {
      return(div(style = box_style("#fff3cd", "#ffe69c", "#664d03"),
                 "The selection changed. Click 'Confirm selection' again."))
    }
    div(style = box_style("#d1e7dd", "#badbcc", "#0f5132"),
        sprintf("Confirmed: %d program(s). Run and the Output & Log checks use exactly these.", length(rv$selected)))
  })
  
  observeEvent(input$run_btn, {
    if (!need_cfg()) return(invisible(NULL))
    # Run exactly the CONFIRMED selection, in table order (= run order).
    if (!need_selection()) return(invisible(NULL))
    to_run <- rv$selected
    
    if (identical(rv$cfg$r_engine_mode, "custom")) {
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
      win_rscript <- rv$cfg$rscript_path
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
      to_win_bs <- function(p) gsub("/", "\\\\", p)
      c_programs_dir <- file.path(hardcoded_path(), PROGRAMS_SUBDIR)
      win_programs <- vapply(to_run, function(p) to_win_bs(file.path(c_programs_dir, substring(p, nchar(programs_dir) + 2L))),
                             character(1))
      win_programs_dir <- to_win_bs(c_programs_dir)
      
      # The programs must still be the originals (their own hardcoded paths intact).
      chk <- scan_for_drive_refs(list_files_by_ext(programs_dir, c("R"), recursive = TRUE),
                                 path_patterns(extract_drive_prefix(hardcoded_path())), ignore_case = FALSE)
      if (sum(chk$total, na.rm = TRUE) == 0) {
        showNotification(paste0("These programs look already modified by an earlier set-up. Run 'Set up testing environment' ",
                                "again with 'Force re-copy' ticked, then run again."), type = "warning", duration = 15)
        add_log("WARNING: the programs look already modified -- run 'Set up testing environment' again with 'Force re-copy'.")
      }
      
      # Working folder on C: -- copy from the shared drive (robocopy: new/changed
      # files only, nothing deleted), run there with the programs' own original
      # paths, then copy output + logs back so this app can see them. The working
      # folder on C: is NOT deleted here -- the server cleans it up itself.
      # robocopy exit codes 0-7 are not proof the files arrived (a failed
      # attribute change on the shared drive can still return < 8), so the copy
      # back is checked file by file.
      src_final_win <- to_windows_path(effective_dest_root(), linux_home, win_drive)
      c_root_win <- to_win_bs(hardcoded_path())
      copy_lines <- c(
        'cat("=== Preparing files ===\\n")',
        sprintf("c_root <- %s", deparse(c_root_win)),
        sprintf("rc <- system2('robocopy', args = c(shQuote(%s), shQuote(c_root), '/E', '/XO', '/R:2', '/W:2', '/NFL', '/NDL', '/NP'))",
                deparse(src_final_win)),
        "if (rc >= 8) stop('could not prepare the working folder (robocopy exit code ', rc, ')')",
        'cat("=== Files ready ===\\n\\n")'
      )
      copy_back_lines <- c(
        'cat("=== Saving results ===\\n")',
        "results_ok <- TRUE",
        sprintf("for (d in c(%s, 'logs')) {", deparse(OUTPUT_SUBDIR)),
        sprintf("  from <- paste0(%s, '\\\\', d)", deparse(c_root_win)),
        sprintf("  to <- paste0(%s, '\\\\', d)", deparse(src_final_win)),
        "  if (!dir.exists(from)) next",
        "  rc <- system2('robocopy', args = c(shQuote(from), shQuote(to), '/E', '/XO', '/COPY:DT', '/DCOPY:T', '/R:2', '/W:2', '/NFL', '/NDL', '/NP'))",
        "  src_files <- list.files(from, recursive = TRUE)",
        "  n_missing <- sum(!file.exists(file.path(to, src_files)))",
        "  cat('  ', d, ': ', length(src_files) - n_missing, ' of ', length(src_files), ' file(s) are on the shared drive (robocopy exit code ', rc, ')\\n', sep = '')",
        "  if (rc >= 8 || n_missing > 0) results_ok <- FALSE",
        "}",
        sprintf("if (!results_ok) cat('Some results could NOT be saved to the shared drive. They are still in', %s, '\\n')", deparse(c_root_win)),
        sprintf("cat('Working files are kept in', %s, '\\n')", deparse(c_root_win))
      )
      c_check_lines <- c(
        "missing_programs <- programs[!file.exists(programs)]",
        "if (length(missing_programs) > 0) {",
        "  cat('ERROR: these programs were not found in the working folder:\\n')",
        "  cat(paste0('  ', missing_programs), sep = '\\n')",
        "  stop('programs missing')",
        "}"
      )
      
      # Dedicated package library: packages never land in the person's C:
      # user library / the R install itself, and the run can't silently reuse
      # packages installed earlier. Lives at the root of the person's drive
      # (outside the copied submission folder, so the path-replace step never
      # touches installed packages).
      # One library PER R minor version (e.g. safe_r_lib_4.5): compiled packages
      # (.dll) built for one R version cannot be loaded by another, so a shared
      # folder breaks as soon as the Rscript path points at a different R. The
      # suffix is worked out at run time from the R that executes the script.
      win_lib <- to_windows_path(file.path(normalize_prefix(linux_home), WIN_LIB_DIRNAME), linux_home, win_drive)
      lib_runtime_expr <- paste0("paste0(", deparse(gsub("\\\\", "/", win_lib)),
                                 ", '_', R.version$major, '.', sub('\\\\..*$', '', R.version$minor))")
      add_log(sprintf("Windows run will use a dedicated package library: %s_<R version>", win_lib))
      lib_env_lines <- c(
        # Forward slashes on purpose: R reports .libPaths() entries that way, and
        # the TFL programs compare them to Sys.getenv("R_LIBS_USER") as plain text.
        paste0("lib_dir <- ", lib_runtime_expr),
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
      # Its code travels inside the run script and is written to a temp file when the
      # run starts, so nothing extra is left on the shared drive.
      boot_text <- bootstrap_script_lines(win_programs_dir,
                                          lib_runtime_expr,
                                          get_package_repos(), win_programs, isTRUE(rv$cfg$use_pinned_versions),
                                          rv$cfg$pins_file, rv$cfg$pin_only, isTRUE(rv$cfg$pin_compiled))
      bootstrap_lines <- c(
        'cat("=== Preparing packages (company packages from programs folder, public from Artifactory) ===\\n")',
        "boot_file <- tempfile(fileext = '.R')",
        paste0("writeLines(", paste(deparse(boot_text), collapse = ""), ", boot_file)"),
        "system2(rscript_exe, args = shQuote(boot_file))",
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
        win_env_r <- to_win_bs(file.path(c_programs_dir, "env.R"))
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
        copy_lines,
        c_check_lines,
        bootstrap_lines,
        check_lines,
        env_setup_lines,
        "Sys.setenv(R_DIAG_TRACE = '1')",
        "for (i in seq_along(programs)) {",
        "  p <- programs[i]",
        '  cat("=== [", i, "/", length(programs), "] Running:", p, "===\\n")',
        "  status <- system2(rscript_exe, args = shQuote(p))",
        '  cat("=== Finished (exit code", status, "):", p, "===\\n\\n")',
        "}",
        copy_back_lines,
        'cat("All programs finished. Press Enter to close this window.\\n")',
        "invisible(readline())"
      )
      # Rscript path with backslashes, for the .bat below.
      win_rscript_display <- gsub("/", "\\\\", win_rscript)
      bat_path <- file.path(effective_dest_root(), "run_selected_programs.bat")
      win_bat_path <- to_windows_path(bat_path, linux_home, win_drive)
      add_log(sprintf("Generated the Windows run script for %d program(s): %s", length(win_programs), win_bat_path))
      
      # ONE file: the batch part on top, the R code below it. The batch part starts
      # Rscript, which reads this same file, skips the batch lines and runs the R code
      # (`exit /b` stops cmd before it ever reaches the R code). Reading it with R
      # avoids any batch escaping of characters such as % | > in the R code.
      r_text <- unlist(strsplit(wrapper_lines, "\n", fixed = TRUE))
      bat_head <- function(n_skip) c(
        "@echo off",
        paste0('set "RSCRIPT=', win_rscript_display, '"'),
        'set "SELF=%~f0"',
        'set "SELF=%SELF:\\=/%"',
        'if not exist "%RSCRIPT%" goto no_rscript',
        paste0('"%RSCRIPT%" -e "eval(parse(text = readLines(\'%SELF%\')[-seq_len(', n_skip, ')]))"'),
        "echo.",
        "echo Finished. Exit code: %errorlevel%",
        "goto done",
        ":no_rscript",
        "echo Rscript not found: %RSCRIPT%",
        ":done",
        "echo.",
        "pause",
        "exit /b",
        ":: ---- R code below: it is read by Rscript, not by this batch file ----")
      bat_lines <- c(bat_head(length(bat_head(0L))), r_text)
      
      # (This app runs on Linux, so it can create the file but cannot start it -- the
      # double-click is the person's.)
      bat_ok <- tryCatch({
        con <- file(bat_path, open = "wb")
        tryCatch(writeLines(bat_lines, con, sep = "\r\n"), finally = close(con))
        TRUE
      }, error = function(e) FALSE)
      # A file created from Linux gets mode 644 (no execute bit); Windows can refuse to
      # start such a file from the shared drive ("You do not have permission").
      # Give it the execute bits (755).
      if (bat_ok) tryCatch(Sys.chmod(bat_path, mode = "0755", use_umask = FALSE),
                           error = function(e) NULL, warning = function(w) NULL)
      add_log(if (bat_ok) {
        sprintf("Created %s -- double-click it to start the run.", win_bat_path)
      } else {
        "ERROR: could not write the .bat file. Check that the working folder is writable."
      })
      
      showNotification(if (bat_ok) sprintf("Script ready. Double-click %s to start the run.", win_bat_path)
                       else "Could not write the .bat file. See the Console Log.",
                       type = if (bat_ok) "message" else "error", duration = 20)
      return(invisible(NULL))
    }
    
    if (!file.exists(effective_rscript_path())) {
      showNotification("Rscript path is not valid. Fix it in Advanced settings first.", type = "error")
      return(invisible(NULL))
    }
    if (isTRUE(rv$cfg$clean_library) && is.null(rv$clean_lib_dir)) {
      rv$clean_lib_dir <- file.path(tempdir(), paste0("clean_rlib_", format(Sys.time(), "%Y%m%d_%H%M%S")))
      dir.create(rv$clean_lib_dir, recursive = TRUE, showWarnings = FALSE)
      add_log("Clean package library set up (packages will install fresh during this run).")
    }
    rv$run_queue <- to_run
    rv$run_results <- list()
    add_log(sprintf("Queued %d program(s) to run%s.", length(rv$run_queue),
                    if (isTRUE(rv$cfg$clean_library)) " with a CLEAN package library" else ""))
    
    if (isTRUE(rv$cfg$clean_library)) {
      # A clean library has zero packages, so first make everything the
      # programs need available: company packages from the programs folder,
      # public ones from the company-configured repo. No user input needed.
      boot_script <- file.path(tempdir(), "bootstrap_packages.R")
      writeLines(bootstrap_script_lines(file.path(effective_dest_root(), PROGRAMS_SUBDIR),
                                        deparse(rv$clean_lib_dir), get_package_repos(), to_run,
                                        isTRUE(rv$cfg$use_pinned_versions),
                                        rv$cfg$pins_file, rv$cfg$pin_only, isTRUE(rv$cfg$pin_compiled)),
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
    win_libs_linux <- Sys.glob(file.path(normalize_prefix(Sys.getenv("HOME")), paste0(WIN_LIB_DIRNAME, "*")))
    win_libs_linux <- win_libs_linux[dir.exists(win_libs_linux)]
    if (length(win_libs_linux) > 0) {
      unlink(win_libs_linux, recursive = TRUE)
      add_log(sprintf("Removed the dedicated Windows package librar%s (%s).",
                      if (length(win_libs_linux) > 1) "ies" else "y", paste(basename(win_libs_linux), collapse = ", ")))
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
      launch <- build_launch_args(next_prog, rv$cfg$clean_library, rv$clean_lib_dir)
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
    if (!need_cfg()) return(invisible(NULL))
    output_dir <- file.path(effective_dest_root(), OUTPUT_SUBDIR)
    source_output_dir <- file.path(source_final_path(), OUTPUT_SUBDIR)
    out_exts <- OUTPUT_EXTENSIONS
    
    if (!need_selection()) return(invisible(NULL))
    # The programs confirmed in "1. Set Up Testing Environment & Run Programs".
    target_programs <- basename(rv$selected)
    
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
      
      server_compare <- if (exists_ok) compare_outputs_to_server(found, rv$cfg$server_output_path) else NA_character_
      
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
    if (!need_cfg()) return(invisible(NULL))
    if (!need_selection()) return(invisible(NULL))
    
    # Strictly use "1. Submission Program Folder"'s own output -- NOT this
    # tool's own local test run. The test run executes under whatever R
    # engine is running this app (e.g. the Linux Posit Workbench engine
    # when accessed via the browser, even through a Windows/Citrix client),
    # which can give numerically different results than a genuine Windows
    # RStudio run (different BLAS/LAPACK, package versions, RNG behavior in
    # multiple-imputation steps, etc.). Comparing the program folder's own
    # pre-existing output against the server avoids that confound.
    source_output_dir <- file.path(source_final_path(), OUTPUT_SUBDIR)
    # EXCEPTION: with the local Windows R engine the test run IS a genuine Windows
    # run, so its own output (in the working copy, incl. subfolders such as
    # output/posthoc) is what gets compared first; the source folder's delivered
    # output is only the fallback.
    work_output_dir <- file.path(effective_dest_root(), OUTPUT_SUBDIR)
    use_work <- identical(rv$cfg$r_engine_mode, "custom")
    server_dir <- rv$cfg$server_output_path
    server_ok <- !is.null(server_dir) && nzchar(trimws(server_dir)) && dir.exists(server_dir)
    
    # Compute the full itemized detail ONCE per program (not twice) -- the
    # on-screen one-line-per-program summary is then just the first row of
    # that same detail, so both views come from a single comparison pass.
    per_program <- lapply(basename(rv$selected), function(prog_name) {
      base <- tools::file_path_sans_ext(basename(prog_name))
      ards_name <- paste0(base, "_ards.csv")
      local_csv <- NA_character_
      local_from <- "not found"
      if (use_work) {
        local_csv <- find_file_shallow(work_output_dir, ards_name)
        if (!is.na(local_csv)) local_from <- "this test run"
      }
      if (is.na(local_csv)) {
        local_csv <- find_file_shallow(source_output_dir, ards_name)
        if (!is.na(local_csv)) local_from <- "delivered output in the source folder"
      }
      if (is.na(local_csv)) local_csv <- file.path(source_output_dir, ards_name)   # keeps the LOCAL_MISSING path
      server_csv <- if (server_ok) find_file_shallow(server_dir, ards_name) else NA_character_
      if (is.na(server_csv)) server_csv <- file.path(if (is.null(server_dir)) "" else server_dir, ards_name)
      add_log(sprintf("ARDS %s: local = %s; server = %s", base, local_from,
                      if (file.exists(server_csv)) "found" else "not found"))
      
      detail <- if (!server_ok) {
        data.frame(row = NA_integer_, column = NA_character_,
                   local_value = NA_character_, server_value = NA_character_,
                   status = "SERVER_PATH_NOT_FOUND", stringsAsFactors = FALSE)
      } else {
        compare_ards_to_server_detail(local_csv, server_csv)
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
    
    n_match <- sum(startsWith(rv$ards_compare_result$result, "MATCH"))
    if (is.na(xlsx_engine())) {
      showNotification(paste0("Neither the 'openxlsx' nor the 'writexl' R package is installed on this server, ",
                              "so the report will download as a CSV instead of an Excel file."),
                       type = "warning", duration = 15)
    }
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
    filename = function() paste0("output_comparison_report_", format(Sys.time(), "%Y%m%d_%H%M%S"),
                                 if (is.na(xlsx_engine())) ".csv" else ".xlsx"),
    content = function(file) {
      req(rv$ards_compare_detail)
      if (is.na(xlsx_engine())) {
        write.csv(rv$ards_compare_detail, file, row.names = FALSE)
      } else {
        write_report_xlsx(build_ards_report_sheets(rv$ards_compare_detail), file)
      }
    }
  )
}

shinyApp(ui, server)