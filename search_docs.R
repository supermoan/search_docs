# search_docs.py
# Andre Moan <andre.moan@hi.no>

# This script searches through the contents of all word documents in a given directory and all
# its subdirectories, looking for matches to a set of user-defined strings. This can be used to
# identify all word documents that contain one of the specified strings. Matching is done
# with regular expressions, so it is quite flexible.

install.packages("textreadr")
library(textreadr)

search_dir <- "/Users/a21997/Documents/Bifangstestimering/Data/pelagisk/"
search_recursive <- TRUE
search_strings <- c("[h|k]val", "kn.l", "spekkh[o|u]gg[a|e]r", "nise", "spring[a|e]r", "[h|k]vitnos", "[h|k]vitskjeving",
    "delfin", "sj.pattedyr", "sel", "steinkobbe", "havert", "gr.nnlands?sel", "ringsel","klappmyss","storkobbe", "hvalross")

files <- list.files(search_dir, pattern = "\\.doc.?", recursive=T, full.names=T)

extension <- function(filename) {
    x <- unlist(strsplit(basename(filename), ".", fixed=T))
    x[length(x)]
}

search <- function(filename) {
    type <- extension(filename)
    if (type == "doc") {
        x <- read_doc(filename)
    } else {
        x <- paste(read_docx(filename), collapse = "")
    }
    
    x <- tolower(x)
    
    sapply(search_strings, grepl, x)
}

matches <- as.data.frame(t(sapply(files, search, USE.NAMES=F)))
pos_matches <- rowSums(matches) > 0

# show results for all files
data.frame(filename = basename(files), match = rowSums(matches))

# show only positive matches
data.frame(filename = basename(files)[pos_matches], match = rowSums(matches)[pos_matches])
