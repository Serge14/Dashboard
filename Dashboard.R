setwd("/home/sergiy/Documents/Work/Nutricia/Rework/201802")

library(data.table)
library(reshape2)
library(googlesheets)

# Read all necessary files

data = fread("BFprocessed.csv", header = TRUE, stringsAsFactors = FALSE, data.table = TRUE)

# Set current month
YTD.No = 2

getTable = function(df) {
    
    nc = length(df)
    
    df$YTDLY = df[, Reduce(`+`, .SD), .SDcols = (nc - YTD.No - 11):(nc - 12)]
    df$YTD = df[, Reduce(`+`, .SD), .SDcols = (nc - YTD.No + 1):nc]
    df$MATLY = df[, Reduce(`+`, .SD), .SDcols = (nc - 23):(nc - 12)]
    df$MAT = df[, Reduce(`+`, .SD), .SDcols = (nc - 11):nc]
    
    df = df[, .SD, .SDcols=c(1, nc-1, nc, (nc+1):(nc+4))]
    
    df = df[, paste0(names(df)[2:length(df)], "_MS") := lapply(.SD, function(x) x/sum(x)), 
            .SDcols = 2:length(df)]          
    
    df = df[, .SD, .SDcols = -c(2:7)]
    
    df$DeltaCMbps = df[, 10000*Reduce(`-`, .SD), .SDcols = c(3, 2)]
    df$DeltaYTDbps = df[, 10000*Reduce(`-`, .SD), .SDcols = c(5, 4)]
    df$DeltaMATbps = df[, 10000*Reduce(`-`, .SD), .SDcols = c(7, 6)]
    
    result = setcolorder(df, c(1, 4:5, 9, 6:7, 10, 2:3, 8))
}

dashTable = function(measure, companiesToShow, brandsToShow, filterSegments = NULL) {

    df = data[eval(parse(text = filterSegments)), .(Items = sum(ITEMSC), Value = sum(VALUEC), Volume = sum(VOLUMEC)),
              by = .(Brand, PS0, PS2, PS3, PS, Company, Ynb, Mnb, PriceSegment)]
    
    companies = data.table::dcast(df, Company~Ynb+Mnb, fun = sum, value.var = measure)
    companies = getTable(companies)
    brands = data.table::dcast(df, Brand~Ynb+Mnb, fun = sum, value.var = measure)
    brands = getTable(brands)
    
    companies=companies[Company %in% companiesToShow]
    brands = brands[Brand %in% brandsToShow]
    names(brands)[1] = "Company"
    result = Reduce("rbind", list(companies, brands))
    
    return(result)
}

gs_title("Dashboard Baby Food Nutricia")
gs_object <- gs_key("1KQIWdgdqH6EuA1YyU4cjfL0ZD1HXrI6LFCmRtEfU2Oc")


#IMF + Dry Food + Puree

companiesToShow = c("NUTRICIA",
                    "NESTLE",
                    "KHOROLSKII MK",
                    "HIPP",
                    "ASSOCIACIYA DP",
                    "VITMARK")

brandsToShow = c("MALYSH ISTR")

result = dashTable("Volume", companiesToShow, brandsToShow, 
                   'PS0 == "IMF" | PS2 == "DRY FOOD" | PS3 == "SAVOURY MEAL" | PS3 == "FRUITS"')

n1 = which(result[,Company == "NUTRICIA"])
n2 = dim(result)[1]

result[7,2:10] = result[5, 2:10]-result[7,2:10]
result[7,4] = (result[7,3]-result[7,2])*10000
result[7,7] = (result[7,6]-result[7,5])*10000
result[7,10] = (result[7,9]-result[7,8])*10000
result[7,1] = "NUTRICIA w/o MALYSH ISTR"
newOrder = c(5, 7, 4, 3, 2, 1, 6)
setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

gs_edit_cells(gs_object, ws="Sheet1", input = result, anchor="O7", col_names=FALSE, trim=FALSE)

result = dashTable("Value", companiesToShow, brandsToShow, 
                   'PS0 == "IMF" | PS2 == "DRY FOOD" | PS3 == "SAVOURY MEAL" | PS3 == "FRUITS"')

result[7,2:10] = result[5,2:10]-result[7,2:10]
result[7,4] = (result[7,3]-result[7,2])*10000
result[7,7] = (result[7,6]-result[7,5])*10000
result[7,10] = (result[7,9]-result[7,8])*10000
result[7,1] = "NUTRICIA w/o MALYSH ISTR"
newOrder = c(5, 7, 4, 3, 2, 1, 6)
setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

gs_edit_cells(gs_object, ws="Sheet1", input = result, anchor="A7", col_names=FALSE, trim=FALSE)


#IMF + Dry Food

companiesToShow = c("NUTRICIA",
                    "NESTLE",
                    "KHOROLSKII MK",
                    "HIPP",
                    "DMK HUMANA",
                    "ASSOCIACIYA DP",
                    "ABBOTT LAB",
                    "DROGA KOLINSKA",
                    "FRIESLAND CAMPINA",
                    "BELLAKT",
                    "EKONIYA",
                    "BIBICALL")

brandsToShow = c("MALYSH ISTR")

result = dashTable("Volume", companiesToShow, brandsToShow, 'PS0 == "IMF" | PS2 == "DRY FOOD"')

result[13,2:10] = result[12,2:10]-result[13,2:10]
result[13,4] = (result[13,3]-result[13,2])*10000
result[13,7] = (result[13,6]-result[13,5])*10000
result[13,10] = (result[13,9]-result[13,8])*10000
result[13,1] = "NUTRICIA w/o MALYSH ISTR"
newOrder = c(12, 13, 11, 10, 2, 9, 3, 6, 5, 8, 1, 7, 4)
setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

gs_edit_cells(gs_object, ws="Sheet1", input = result, anchor="O18", col_names=FALSE, trim=FALSE)

result = dashTable("Value", companiesToShow, brandsToShow, 'PS0 == "IMF" | PS2 == "DRY FOOD"')

result[13,2:10] = result[12,2:10]-result[13,2:10]
result[13,4] = (result[13,3]-result[13,2])*10000
result[13,7] = (result[13,6]-result[13,5])*10000
result[13,10] = (result[13,9]-result[13,8])*10000
result[13,1] = "NUTRICIA w/o MALYSH ISTR"
newOrder = c(12, 13, 11, 10, 2, 9, 3, 6, 5, 8, 1, 7, 4)
setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

gs_edit_cells(gs_object, ws="Sheet1", input = result, anchor="A18", col_names=FALSE, trim=FALSE)

# IMF
companiesToShow = c("NUTRICIA",
                    "NESTLE",
                    "KHOROLSKII MK",
                    "HIPP",
                    "DMK HUMANA",
                    "FRIESLAND CAMPINA",
                    "BELLAKT")

brandsToShow = c("NUTRILON",
                 "MILUPA",
                 "MALYSH ISTR",
                 "NAN",
                 "NESTOGEN",
                 "MALYSH KH",
                 "MALYUTKA KH",
                 "MOLOKO")

result = dashTable("Volume", companiesToShow, brandsToShow, 'PS0 == "IMF"')
newOrder = c(13, 14, 15, 12, 11, 7, 4, 2, 8, 9, 3, 10, 5, 6, 1)
setorder(result[, .r := newOrder], .r)[, .r := NULL]

gs_edit_cells(gs_object, ws="Sheet1", input = result, anchor="O35", col_names=FALSE, trim=FALSE)

result = dashTable("Value", companiesToShow, brandsToShow, 'PS0 == "IMF"')
newOrder = c(13, 14, 15, 12, 11, 7, 4, 2, 8, 9, 3, 10, 5, 6, 1)
setorder(result[, .r := newOrder], .r)[, .r := NULL]

gs_edit_cells(gs_object, ws="Sheet1", input = result, anchor="A35", col_names=FALSE, trim=FALSE)


# IF Base
companiesToShow = c("NUTRICIA",
                    "NESTLE",
                    "KHOROLSKII MK",
                    "HIPP",
                    "DMK HUMANA",
                    "FRIESLAND CAMPINA",
                    "BELLAKT")

brandsToShow = c("NUTRILON",
                 "MILUPA",
                 "MALYSH ISTR",
                 "NAN",
                 "NESTOGEN",
                 "MALYUTKA KH",
                 "MOLOKO")

result = dashTable("Volume", companiesToShow, brandsToShow, 'PS2 == "IF" & PS3 == "BASE"')
newOrder = c(12, 13, 14, 11, 10, 7, 4, 2, 8, 3, 9, 5, 6, 1)
setorder(result[, .r := newOrder], .r)[, .r := NULL]

gs_edit_cells(gs_object, ws="Sheet1", input = result, anchor="O54", col_names=FALSE, trim=FALSE)

result = dashTable("Value", companiesToShow, brandsToShow, 'PS2 == "IF" & PS3 == "BASE"')
newOrder = c(12, 13, 14, 11, 10, 7, 4, 2, 8, 3, 9, 5, 6, 1)
setorder(result[, .r := newOrder], .r)[, .r := NULL]

gs_edit_cells(gs_object, ws="Sheet1", input = result, anchor="A54", col_names=FALSE, trim=FALSE)

# FO Base
companiesToShow = c("NUTRICIA",
                    "NESTLE",
                    "KHOROLSKII MK",
                    "HIPP",
                    "DMK HUMANA",
                    "FRIESLAND CAMPINA",
                    "BELLAKT")

brandsToShow = c("NUTRILON",
                 "MILUPA",
                 "MALYSH ISTR",
                 "NAN",
                 "NESTOGEN",
                 #"NESTLE",
                 #"MALYSH KH",
                 "MALYUTKA KH")

result = dashTable("Volume", companiesToShow, brandsToShow, 'PS2 == "FO" & PS3 == "BASE"')
newOrder = c(11, 12, 13, 10, 9, 7, 4, 2, 8, 3, 5, 6, 1)
setorder(result[, .r := newOrder], .r)[, .r := NULL]

gs_edit_cells(gs_object, ws="Sheet1", input = result, anchor="O72", col_names=FALSE, trim=FALSE)

result = dashTable("Value", companiesToShow, brandsToShow, 'PS2 == "FO" & PS3 == "BASE"')
newOrder = c(11, 12, 13, 10, 9, 7, 4, 2, 8, 3, 5, 6, 1)
setorder(result[, .r := newOrder], .r)[, .r := NULL]

gs_edit_cells(gs_object, ws="Sheet1", input = result, anchor="A72", col_names=FALSE, trim=FALSE)

# GUM Base
companiesToShow = c("NUTRICIA",
                    "NESTLE",
                    "KHOROLSKII MK",
                    "DMK HUMANA",
                    "FRIESLAND CAMPINA")

brandsToShow = c("NUTRILON",
                 "MILUPA",
                 "MALYSH ISTR",
                 "NAN",
                 "NESTOGEN",
                 "MALYUTKA KH")

result = dashTable("Volume", companiesToShow, brandsToShow, 'PS2 == "GUM" & PS3 == "BASE"')
newOrder = c(11, 10, 9, 7, 4, 2, 8, 3, 5, 6, 1)
setorder(result[, .r := newOrder], .r)[, .r := NULL]

gs_edit_cells(gs_object, ws="Sheet1", input = result, anchor="O89", col_names=FALSE, trim=FALSE)

result = dashTable("Value", companiesToShow, brandsToShow, 'PS2 == "GUM" & PS3 == "BASE"')
newOrder = c(11, 10, 9, 7, 4, 2, 8, 3, 5, 6, 1)
setorder(result[, .r := newOrder], .r)[, .r := NULL]

gs_edit_cells(gs_object, ws="Sheet1", input = result, anchor="A89", col_names=FALSE, trim=FALSE)

# Specials
companiesToShow = c("NUTRICIA",
                    "NESTLE",
                    "HIPP",
                    "DMK HUMANA",
                    "FRIESLAND CAMPINA",
                    "BIBICALL",
                    "ABBOTT LAB")

brandsToShow = c("")

result = dashTable("Volume", companiesToShow, brandsToShow, 'PS3 == "SPECIALS"')
newOrder = c(6, 5, 4, 3, 7, 2, 1)
setorder(result[, .r := newOrder], .r)[, .r := NULL]

gs_edit_cells(gs_object, ws="Sheet1", input = result, anchor="O104", col_names=FALSE, trim=FALSE)

result = dashTable("Value", companiesToShow, brandsToShow, 'PS3 == "SPECIALS"')
newOrder = c(6, 5, 4, 3, 7, 2, 1)
setorder(result[, .r := newOrder], .r)[, .r := NULL]

gs_edit_cells(gs_object, ws="Sheet1", input = result, anchor="A104", col_names=FALSE, trim=FALSE)

# Dry Food
companiesToShow = c("NUTRICIA",
                    "NESTLE",
                    "KHOROLSKII MK",
                    "DROGA KOLINSKA",
                    "HEINZ",
                    "ASSOCIACIYA DP",
                    "HIPP",
                    "DMK HUMANA",
                    "EKSTREYD",
                    "BELLAKT")

brandsToShow = c("NUTRILON",
                 "MILUPA")

result = dashTable("Volume", companiesToShow, brandsToShow, 'PS2 == "DRY FOOD"')
newOrder = c(5, 10, 11, 6, 9, 7, 8, 12, 4, 3, 2, 1)
setorder(result[, .r := newOrder], .r)[, .r := NULL]

gs_edit_cells(gs_object, ws="Sheet1", input = result, anchor="O115", col_names=FALSE, trim=FALSE)

result = dashTable("Value", companiesToShow, brandsToShow, 'PS2 == "DRY FOOD"')
newOrder = c(5, 10, 11, 6, 9, 7, 8, 12, 4, 3, 2, 1)
setorder(result[, .r := newOrder], .r)[, .r := NULL]

gs_edit_cells(gs_object, ws="Sheet1", input = result, anchor="A115", col_names=FALSE, trim=FALSE)

# Cereals
companiesToShow = c("NUTRICIA",
                    "NESTLE",
                    "KHOROLSKII MK",
                    "DROGA KOLINSKA",
                    "ASSOCIACIYA DP",
                    "HIPP",
                    "DMK HUMANA",
                    "EKSTREYD",
                    "BELLAKT")

brandsToShow = c("NUTRILON",
                 "MILUPA")

result = dashTable("Volume", companiesToShow, brandsToShow, 
#                    'PS3 == "INSTANT CEREALS" | PS3 == "READY TO EAT CEREALS"')
                    'PS3 == "INSTANT CEREALS"')
newOrder = c(5, 9, 10, 6, 7, 8, 11, 4, 3, 2, 1)
setorder(result[, .r := newOrder], .r)[, .r := NULL]

gs_edit_cells(gs_object, ws="Sheet1", input = result, anchor="O131", col_names=FALSE, trim=FALSE)

result = dashTable("Value", companiesToShow, brandsToShow, 
#                   'PS3 == "INSTANT CEREALS" | PS3 == "READY TO EAT CEREALS"')
                    'PS3 == "INSTANT CEREALS"')
newOrder = c(5, 9, 10, 6, 7, 8, 11, 4, 3, 2, 1)
setorder(result[, .r := newOrder], .r)[, .r := NULL]

gs_edit_cells(gs_object, ws="Sheet1", input = result, anchor="A131", col_names=FALSE, trim=FALSE)

# Cereal Biscuits
companiesToShow = c("NUTRICIA",
                    "HEINZ",
                    "EKONIYA",
                    "DROGA KOLINSKA",
                    "HIPP",
                    "LAGODA")

brandsToShow = c("")

result = dashTable("Volume", companiesToShow, brandsToShow, 'PS3 == "CEREAL BISCUITS"')
newOrder = c(5, 2, 3, 6, 4, 1)
setorder(result[, .r := newOrder], .r)[, .r := NULL]

gs_edit_cells(gs_object, ws="Sheet1", input = result, anchor="O146", col_names=FALSE, trim=FALSE)

result = dashTable("Value", companiesToShow, brandsToShow, 'PS3 == "CEREAL BISCUITS"')
newOrder = c(5, 2, 3, 6, 4, 1)
setorder(result[, .r := newOrder], .r)[, .r := NULL]

gs_edit_cells(gs_object, ws="Sheet1", input = result, anchor="A146", col_names=FALSE, trim=FALSE)

# Puree

companiesToShow = c("ASSOCIACIYA DP",
                    "EKONIYA",
                    "HAME",
                    "HIPP",
                    "NESTLE",
                    "NUTRICIA",
                    "VITMARK")

brandsToShow = c("FRUTA PYURESHKA",
                  "HAMANEK",
                  "HAME",
                  "BEBIVITA",
                  "HIPP")

result = dashTable("Volume", companiesToShow, brandsToShow, 
                   'PS3 == "FRUITS" | PS3 == "SAVOURY MEAL"')
#newOrder = c(3, 2, 1, 12, 5, 7, 6, 8, 9, 10, 11, 4)
newOrder = c(6, 5, 7, 1, 8, 12, 4, 2, 9, 10, 11, 3)
setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

gs_edit_cells(gs_object, ws="Sheet1", input = result, anchor="O156", col_names=FALSE, trim=FALSE)

result = dashTable("Value", companiesToShow, brandsToShow, 
                   'PS3 == "FRUITS" | PS3 == "SAVOURY MEAL"')
newOrder = c(6, 5, 7, 1, 8, 12, 4, 2, 9, 10, 11, 3)
setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

gs_edit_cells(gs_object, ws="Sheet1", input = result, anchor="A156", col_names=FALSE, trim=FALSE)
