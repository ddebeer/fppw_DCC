rm_substrings <- function(character_string, substrings_to_rm){
  for(substring in substrings_to_rm){
    character_string <- gsub(substring, "", character_string, fixed = TRUE)
  }
  character_string
}



add_blocks1 <- function(rdocx, list_of_block_lists){
  for(block_list in list_of_block_lists){
    rdocx <- body_add_blocks(rdocx, block_list)
  }
  rdocx
}

add_blocks2 <- function(rdocx, named_list_of_list_of_block_lists){
  names <- names(named_list_of_list_of_block_lists)
  for(name in names){
    rdocx <- body_add_fpar(rdocx, fpar(ftext(name, 
                                             fp_text(font.size = 11, bold = TRUE))))
    rdocx <- add_blocks1(rdocx, named_list_of_list_of_block_lists[[name]])
  }
  rdocx

}



rename_cols <- function(data) {
  dat_names <- names(data)
  dat_names[2:3] <- c("Naam_Std", "Voornaam_Std")
  
  dat_names[21:22] <- c("Status_inschr", "Datum_status_inschr")
  dat_names[24:25] <- c("Status_doc_geg", "Datum_doc_geg")
  dat_names[26:27] <- c("Status_kand_geg", "Datum_kand_geg")
  dat_names[28:29] <- c("Status_BVA", "Datum BVA")
  
  functions <- c("Adm_Prom", paste0("Begel_", 1:5))
  colnames_functions <- c("Aanspreekvorm", "Naam", "Voornaam", 
                          "Functietype", "Vakgroepcode", "Vakgroep")
  
  for(colname in colnames_functions){
    if(colname == "Functietype") {
      dat_names[grep(paste0(colname, "."), dat_names, fixed = TRUE)] <- paste0(colname, "_", functions[-1])
    }
    else {dat_names[grep(paste0(colname, "."), dat_names, fixed = TRUE)] <- paste0(colname, "_", functions)}
    invisible(NULL)
  }
  
  dat_names <- rm_substrings(gsub(" ", "_", dat_names, fixed = TRUE), c(".", "(", ")"))
  names(data) <- dat_names
  data
}



preprocess_data <- function(data){
  data$Basis <- rm_substrings(data$Basis_van_aanvaarding, 
                              c("DiplomaContract ", 
                                "Master of Science in de psychologie: Clinical Psychology, ",
                                ", Master of Science in de psychologie: Clinical Psychology",
                                "Master of Science in de psychologie: Theoretical and Experimental Psychology, ",
                                ", Master of Science in de psychologie: Theoretical and Experimental Psychology",
                                "Master of Science in de pedagogische wetenschappen: Special Education, Disability Studies and Behavioral Disorders, ",
                                ", Master of Science in de pedagogische wetenschappen: Special Education, Disability Studies and Behavioral Disorders",                                "Master of Science in de pedagogische wetenschappen: Pedagogy and Educational Sciences, ", 
                                ", Master of Science in de pedagogische wetenschappen: Pedagogy and Educational Sciences",
                                "Master of Arts in de taal- en letterkunde: French - English, "))
  data[is.na(data)] <- ""
  data
}


extract_info_row <- function(row, type = c("html", ".docx")){
  
  type <- match.arg(type)
  # If more than two promoters, than add ...
  n_promo <- with(row, sum(c(Functietype_Begel_1, Functietype_Begel_2, Functietype_Begel_3, 
                             Functietype_Begel_4, Functietype_Begel_5) == "Promotor", 
                           na.rm = TRUE) + 1)
  add <- switch(as.character(n_promo), 
                "3" = "(motivering drie promotoren)",
                "4" = "(motivering vier promotoren)",
                "5" = "(motivering vijf promotoren)", 
                "")
  
  
  if(type == "html"){
    out <- with(row, paste0("<div><b>", Naam_Std, ", ", Voornaam_Std, "</b>, ", Basis, 
                            " (administratief promotor: ", Voornaam_Adm_Prom, " ", Naam_Adm_Prom, 
                            " (", Vakgroepcode_Adm_Prom, ")), <b>VOEG BIJLAGES TOE</b>. ", add, 
                            "</div> <br> <div><em>Nederlandstalige titel</em>: ", Titel_doct_onderzoek_nl, 
                            ".</div> <br> <div><em>Engelstalige titel</em>: ", Titel_doct_onderzoek_en, ".</div> <br>  <div>",
                            "Begeleidingscommissie: ", Voornaam_Adm_Prom, " ", Naam_Adm_Prom,
                            " (administratief promotor), ", 
                            Voornaam_Begel_1, " ", Naam_Begel_1," (", Functietype_Begel_1, "), ",
                            Voornaam_Begel_2, " ", Naam_Begel_2," (", Functietype_Begel_2, "), ",
                            Voornaam_Begel_3, " ", Naam_Begel_3," (", Functietype_Begel_3, "), ",
                            Voornaam_Begel_4, " ", Naam_Begel_4," (", Functietype_Begel_4, "), ",
                            Voornaam_Begel_5, " ", Naam_Begel_5," (", Functietype_Begel_5, "). </div> <br> <br> "))
    
    out <- rm_substrings(out, c(",   ()", ".."))
    return(out)
  }
  # most of the used functions come from the officer-package
  
  # formatting info
  norm <- fp_text(font.size = 11)
  bold <- fp_text(font.size = 11, bold = TRUE)
  emph <- fp_text(font.size = 11, italic = TRUE)
  para <- fp_par(text.align = "justify", padding.bottom = 0)
  
  
  block <- with(row, block_list(
    fpar(
      ftext(paste0(Naam_Std, ", ", Voornaam_Std, ", "), bold),
      ftext(paste0(Basis, " (administratief promotor: ", Voornaam_Adm_Prom, " ", 
                   Naam_Adm_Prom, " (", Vakgroepcode_Adm_Prom, 
                   ")), VOEG BIJLAGES TOE. ", add), norm),
      fp_p = para),
    fpar(
      ftext("", norm),
      fp_p = para), 
    fpar(
      ftext("Nederlandstalige titel", emph),
      ftext(paste0(": ", Titel_doct_onderzoek_nl), norm),
      fp_p = para),
    fpar(
      ftext("", norm),
      fp_p = para),
    fpar(
      ftext("Engelstalige titel", emph),
      ftext(paste0(": ", Titel_doct_onderzoek_en), norm),
      fp_p = para),
    fpar(
      ftext("", norm),
      fp_p = para),
    fpar(
      ftext(rm_substrings(
        paste0("Begeleidingscommissie: ", Voornaam_Adm_Prom, " ", 
               Naam_Adm_Prom, " (administratief promotor), ", 
               Voornaam_Begel_1, " ", Naam_Begel_1," (", Functietype_Begel_1, "), ",
               Voornaam_Begel_2, " ", Naam_Begel_2," (", Functietype_Begel_2, "), ",
               Voornaam_Begel_3, " ", Naam_Begel_3," (", Functietype_Begel_3, "), ",
               Voornaam_Begel_4, " ", Naam_Begel_4," (", Functietype_Begel_4, "), ",
               Voornaam_Begel_5, " ", Naam_Begel_5," (", Functietype_Begel_5, ")."),
        c(",   ()", "..")), norm),
      fp_p = para),
    fpar(
      ftext("", norm),
      fp_p = para)))
  
  block
}


extract_info_subset <- function(subset, type = c("html", ".docx")){
  type <- match.arg(type)
  if(type == "html"){
    return(do.call(paste0, lapply(seq_len(nrow(subset)), function(row_nr) 
    {extract_info_row(subset[row_nr,])})))
  }
  lapply(seq_len(nrow(subset)), function(row_nr) 
  {extract_info_row(subset[row_nr,], type = type)})
}


extract_info <- function(data, type = c("html", ".docx")){
  type <- match.arg(type)
  dat_split <- split(data, f = data$Beoogde_doctorstitel)
  
  if(type == "html"){
    out_split <- lapply(dat_split, extract_info_subset)
    return(do.call(paste0, lapply(names(out_split), function(name){
      paste0("<h3>", name, "</h3> <br>", out_split[[name]])
      })))
  }

  out_split <- lapply(dat_split, extract_info_subset, type = type)
  out <- read_docx()
  out <- add_blocks2(out, out_split)
  out
  
}
