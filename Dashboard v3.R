setwd("/home/sergiy/Documents/Work/Nutricia/Data/202002")

library(data.table)
# library(reshape2)
# library(googlesheets)
library(openxlsx)

# Read all necessary files

df = fread("df.csv", header = TRUE, stringsAsFactors = FALSE, data.table = TRUE)

df = df[, .(Volume = sum(VolumeC), Value = sum(ValueC)), 
        by = .(PS0, PS2, PS3, PS, Form, Package, Company, Brand, Ynb, Mnb)]

df = df[!(PS0 == "IMF" & Form == "Liquid")]

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

dashTable = function(df, measure, companies.to.show, 
                     brands.to.show, segments.to.filter = NULL) {

    df = df[eval(parse(text = segments.to.filter)), 
              .(Volume = sum(Volume), Value = sum(Value)),
              by = .(PS0, PS2, PS3, Company, Brand, Ynb, Mnb)]
    
    companies = dcast.data.table(df, 
                                 Company ~ Ynb + Mnb, 
                                 fun = sum, 
                                 value.var = measure)
    companies = getTable(companies)
    companies = companies[Company %in% companies.to.show]
    
    brands = dcast.data.table(df, 
                              Brand ~ Ynb + Mnb, 
                              fun = sum, 
                              value.var = measure)
    brands = getTable(brands)
    brands = brands[Brand %in% brands.to.show]
    
    names(brands)[1] = "Company"
    result = rbindlist(list(companies, brands))
    
    rm(list = c("companies", "brands"))
    
    return(result)
}

add.nutricia.wo.malysh = function(result, keep.malysh){
    
    n1 = which(result[, Company == "Malysh Istr"])
    n2 = which(result[, Company == "Nutricia"])
    
    if (keep.malysh == TRUE){malysh.istr = result[n1,]}

    result[n1, 2:10] = result[n2, 2:10] - result[n1, 2:10]
    result[n1, 4] = (result[n1, 3] - result[n1, 2])*10000
    result[n1, 7] = (result[n1, 6] - result[n1, 5])*10000
    result[n1, 10] = (result[n1, 9] - result[n1, 8])*10000
    result[n1, 1] = "Nutricia w/o Malysh Istr"
    
    if (keep.malysh == TRUE){result = rbindlist(list(result, malysh.istr))}
    
    return(result) 
}


wb <- createWorkbook()
addWorksheet(wb, "Sheet1")

#IMF + Dry Food + Puree

companies.to.show = c("Nutricia",
                    "Nestle",
                    "Khorolskii Mk",
                    "Hipp",
                    "Vitmark",
                    "Dmk Humana",
                    "ADP")
brands.to.show = c("Malysh Istr",
                   "Nutrilon",
                   "Milupa",
                   "Nan",
                   "Nestogen",
                   "Nestle",
                   "Gerber",
                   "Malyutka Kh",
                   "Malysh Kh",
                   "Malyshka Kh",
                   "Hipp",
                   "Bebivita")

segments.to.filter = 'PS0 == "IMF" | PS2 == "Dry Food" | 
                   PS3 == "Savoury Meal" | PS3 == "Fruits"'

result = dashTable(df,
                   "Volume",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

result = add.nutricia.wo.malysh(result, TRUE)

newOrder = c(6, 11, 19, 15, 20, 5, 16, 18, 17, 9, 4, 14, 12, 13, 3, 10, 8, 7, 2, 1)

setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

writeData(wb, 
          "Sheet1", 
          "IMF + Dry Food + Puree", 
          rowNames = FALSE, 
          xy = c("O", 35))

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("O", 37), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight1")

result = dashTable(df,
                   "Value",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

result = add.nutricia.wo.malysh(result, TRUE)

setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

# gs_edit_cells(gs_object, ws="Sheet1", input = result, anchor="A7", col_names=FALSE, trim=FALSE)
writeData(wb, 
          "Sheet1", 
          "IMF + Dry Food + Puree", 
          rowNames = FALSE, 
          xy = c("A", 35))

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("A", 37), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

#IMF + Dry Food

companies.to.show = c("Nutricia",
                      "Khorolskii Mk",
                      "Nestle",
                      "Hipp",
                      "Dmk Humana",
                      "Friesland Campina",
                      "Bellakt",
                      "Droga Kolinska",
                      "ADP",
                      "Bibicall",
                      "Ausnutria Group",
                      "Abbott Lab",
                      "Ekoniya")

brands.to.show = c("Malysh Istr",
                   "Nutrilon",
                   "Milupa",
                   "Nan",
                   "Nestogen",
                   "Nestle",
                   "Gerber",
                   "Malyutka Kh",
                   "Malysh Kh",
                   "Malyshka Kh",
                   "Hipp")

segments.to.filter = 'PS0 == "IMF" | PS2 == "Dry Food"'
    
result = dashTable(df,
                   "Volume",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

result = add.nutricia.wo.malysh(result, TRUE)

newOrder = c(13, 16, 24, 20, 25, 12, 21, 23, 22, 14, 11, 19, 17, 18, 10, 15,
             6, 9, 4, 7, 1, 5, 3, 2, 8)
setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("O", 62), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

result = dashTable(df,
                   "Value",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

result = add.nutricia.wo.malysh(result, TRUE)

setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("A", 62), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")


# IMF
companies.to.show = c("Khorolskii Mk",
                      "Nestle",
                      "Nutricia",
                      "Hipp",
                      "Bellakt",
                      "Dmk Humana",
                      "Friesland Campina")

brands.to.show = c("Nutrilon",
                 "Milupa",
                 "Malysh Istr",
                 "Nan",
                 "Nestogen",
                 "Malysh Kh",
                 "Malyutka Kh")
segments.to.filter = 'PS0 == "IMF"'


result = dashTable(df,
                   "Volume",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

newOrder = c(14, 12, 13, 11, 8, 5, 1, 4, 10, 9, 3, 7, 6, 2)
setorder(result[, .r := newOrder], .r)[, .r := NULL]

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("O", 92), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")


result = dashTable(df,
                   "Value",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

setorder(result[, .r := newOrder], .r)[, .r := NULL]

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("A", 92), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

# IF Base
companies.to.show = c("Khorolskii Mk",
                      "Nestle",
                      "Nutricia",
                      "Hipp",
                      "Bellakt",
                      "Dmk Humana",
                      "Friesland Campina")

brands.to.show = c("Nutrilon",
                   "Milupa",
                   "Malysh Istr",
                   "Nan",
                   "Nestogen",
                   "Malyutka Kh")
segments.to.filter = 'PS2 == "IF" & PS3 == "Base"'

result = dashTable(df,
                   "Volume",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)
newOrder = c(11, 12, 13, 10, 8, 5, 1, 4, 9, 3, 7, 6, 2)
setorder(result[, .r := newOrder], .r)[, .r := NULL]

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("O", 111), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

result = dashTable(df,
                   "Value",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

setorder(result[, .r := newOrder], .r)[, .r := NULL]

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("A", 111), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

# FO Base
companies.to.show = c("Khorolskii Mk",
                      "Nestle",
                      "Nutricia",
                      "Hipp",
                      "Bellakt",
                      "Dmk Humana",
                      "Friesland Campina")

brands.to.show = c("Nutrilon",
                 "Milupa",
                 "Malysh Istr",
                 "Nan",
                 "Nestogen",
                 "Malyutka Kh")
segments.to.filter = 'PS2 == "FO" & PS3 == "Base"'

result = dashTable(df, "Volume", companies.to.show, brands.to.show, segments.to.filter)
newOrder = c(13, 11, 12, 10, 8, 5, 1, 4, 9, 3, 7, 6, 2)
setorder(result[, .r := newOrder], .r)[, .r := NULL]

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("O", 129), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")


result = dashTable(df, "Value", companies.to.show, brands.to.show, segments.to.filter)

setorder(result[, .r := newOrder], .r)[, .r := NULL]

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("A", 129), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")
# GUM Base
companies.to.show = c("Khorolskii Mk",
                      "Nestle",
                      "Nutricia",
                      "Bellakt",
                      "Dmk Humana",
                      "Friesland Campina")

brands.to.show = c("Nutrilon",
                 "Milupa",
                 "Malysh Istr",
                 "Nan",
                 "Nestogen",
                 "Malyutka Kh")
segments.to.filter = 'PS2 == "Gum" & PS3 == "Base"'


result = dashTable(df, "Volume", companies.to.show, brands.to.show, segments.to.filter)
newOrder = c(12, 11, 10, 8, 5, 1, 4, 9, 3, 7, 6, 2)
setorder(result[, .r := newOrder], .r)[, .r := NULL]

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("O", 147), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

result = dashTable(df, "Value", companies.to.show, brands.to.show, segments.to.filter)

setorder(result[, .r := newOrder], .r)[, .r := NULL]

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("A", 147), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

# Specials
companies.to.show = c("Nutricia",
                      "Nestle",
                      "Bellakt",
                      "Friesland Campina",
                      "Dmk Humana",
                      "Bibicall",
                      "Ausnutria Group",
                      "Hipp",
                      "Abbott Lab")

brands.to.show = c("")
segments.to.filter = 'PS3 == "Specials"'

result = dashTable(df,
                   "Volume",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)
newOrder = c(9, 6, 7, 4, 3, 5, 8, 2, 1)
setorder(result[, .r := newOrder], .r)[, .r := NULL]

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("O", 164), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

result = dashTable(df,
                   "Value",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

setorder(result[, .r := newOrder], .r)[, .r := NULL]

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("A", 164), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

# Dry Food
companies.to.show = c("Nutricia",
                      "ADP",
                      "Droga Kolinska",
                      "Khorolskii Mk",
                      "Hipp",
                      "Nestle",
                      "Ekoniya",
                      "Bellakt",
                      "Kraft Heinz",
                      "Dmk Humana")

brands.to.show = c("Nutrilon",
                 "Milupa")
segments.to.filter = 'PS2 == "Dry Food"'

result = dashTable(df, "Volume", companies.to.show, brands.to.show, segments.to.filter)
newOrder = c(7, 10, 11, 5, 9, 6, 8, 12, 4, 1, 3, 2)
setorder(result[, .r := newOrder], .r)[, .r := NULL]

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("O", 178), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

result = dashTable(df,
                   "Value",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

setorder(result[, .r := newOrder], .r)[, .r := NULL]

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("A", 178), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")


# Cereals
companies.to.show = c(
    "Nutricia",
    "ADP",
    "Droga Kolinska",
    "Khorolskii Mk",
    "Hipp",
    "Nestle",
    "Bellakt",
    "Dmk Humana",
    "Ekoniya",
    "Ekstreyd"
)

brands.to.show = c("Nutrilon",
                   "Milupa")
segments.to.filter = 'PS3 == "Instant Cereals"'


result = dashTable(df, "Volume", companies.to.show, brands.to.show, segments.to.filter)
newOrder = c(7, 9, 10, 5, 12, 11, 6, 8, 4, 1, 3, 2)
setorder(result[, .r := newOrder], .r)[, .r := NULL]

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("O", 195), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")


result = dashTable(df, "Value", companies.to.show, brands.to.show, segments.to.filter)

setorder(result[, .r := newOrder], .r)[, .r := NULL]

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("A", 195), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")


# Cereal Biscuits
companies.to.show = c(
    "Nutricia",
    "Ekoniya",
    "Kraft Heinz",
    "Lagoda",
    "Droga Kolinska",
    "Fleur Alpine Uab",
    "Hipp"
)

brands.to.show = c("")
segments.to.filter = 'PS3 == "Cereal Biscuits"'

result = dashTable(df, "Volume", companies.to.show, brands.to.show, segments.to.filter)
newOrder = c(6, 2, 4, 7, 3, 5, 1)
setorder(result[, .r := newOrder], .r)[, .r := NULL]

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("O", 212), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")


result = dashTable(df, "Value", companies.to.show, brands.to.show, segments.to.filter)

setorder(result[, .r := newOrder], .r)[, .r := NULL]

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("A", 212), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

# Puree

companies.to.show = c(
    "Vitmark",
    "Nestle",
    "ADP",
    "Hipp",
    "Hame",
    "Nutricia",
    "Ekoniya",
    "Profitzernomarket",
    "Belfood"
)

brands.to.show = c("")
segments.to.filter = 'PS3 == "Fruits" | PS3 == "Savoury Meal"'
    
result = dashTable(df,
                   "Volume",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

newOrder = c(7, 6, 9, 5, 4, 1, 8, 3, 2)

setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("O", 224), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

result = dashTable(df,
                   "Value",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("A", 224), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

saveWorkbook(wb, file = "dashboard.xlsx", overwrite = TRUE)

# Meat Meal

companies.to.show = c("Hame",
                      "Hipp",
                      "Nutricia",
                      "Nestle",
                      "Profitzernomarket",
                      "ADP")

brands.to.show = c("")
segments.to.filter = 'PS == "Meat Meal"'

result = dashTable(df,
                   "Volume",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

newOrder = c(5, 2,3, 4, 6, 1)

setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("O", 238), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

result = dashTable(df,
                   "Value",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("A", 238), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

saveWorkbook(wb, file = "dashboard.xlsx", overwrite = TRUE)

# Veggie Meal

companies.to.show = c(
    "Nestle",
    "Hipp",
    "Nutricia",
    "Profitzernomarket",
    "Belfood",
    "Oasis Group",
    "ADP",
    "Ekoniya"
)

brands.to.show = c("")
segments.to.filter = 'PS == "Veggie Meal"'

result = dashTable(df,
                   "Volume",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

newOrder = c(6, 5, 4, 8, 2, 7, 3, 1)

setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("O", 252), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

result = dashTable(df,
                   "Value",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("A", 252), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

saveWorkbook(wb, file = "dashboard.xlsx", overwrite = TRUE)

# Fruit Jar/Glass

companies.to.show = c(
    "Vitmark",
    "ADP",
    "Nestle",
    "Ekoniya",
    "Nutricia",
    "Hipp",
    "Hame",
    "Belfood",
    "Oasis Group"
)

brands.to.show = c("")
segments.to.filter = 'PS3 == "Fruits" & Package == "Glass"'

result = dashTable(df,
                   "Volume",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

newOrder = c(7, 9, 1, 6, 3, 4, 5, 2, 8)

setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("O", 266), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

result = dashTable(df,
                   "Value",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("A", 266), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

saveWorkbook(wb, file = "dashboard.xlsx", overwrite = TRUE)

# Fruit Pouch

companies.to.show = c("Vitmark",
                      "Nestle",
                      "Hipp",
                      "Hame",
                      "Nutricia",
                      "Profitzernomarket",
                      "Belfood")

brands.to.show = c("")
segments.to.filter = 'PS3 == "Fruits" & Package == "Pouch"'

result = dashTable(df,
                   "Volume",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

newOrder = c(5, 7, 4, 2, 3, 6, 1)

setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("O", 280), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

result = dashTable(df,
                   "Value",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("A", 280), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

saveWorkbook(wb, file = "dashboard.xlsx", overwrite = TRUE)

# AMN

companies.to.show = c("Nutricia",
"Nestle",
"Abbott Lab",
"B. Braun Melsungen",
"Fresenius Kabi")

brands.to.show = c(
    "Nutridrink",
    "Nutrison",
    "P-Am",
    "Pku Nutri",
    "Resource",
    "Peptamen",
    "Modulen",
    "Isosource",
    "Pediasure",
    "Nutricomp",
    "Fresubin"
)

segments.to.filter = 'PS0 == "AMN"'

result = dashTable(df,
                   "Volume",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

newOrder = c(5, 10, 11, 12, 15, 4, 16, 14, 8, 7, 1, 13, 2, 9, 3, 6)

setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("O", 294), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

result = dashTable(df,
                   "Value",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("A", 294), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

saveWorkbook(wb, file = "dashboard.xlsx", overwrite = TRUE)

# Recovery

companies.to.show = c("Nutricia",
                     "B. Braun Melsungen")

brands.to.show = c(
    "Nutridrink",
    "Nutricomp"
)

segments.to.filter = 'PS == "Recovery"'

result = dashTable(df,
                   "Volume",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

newOrder = c(2, 4, 1, 3)

setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("O", 315), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

result = dashTable(df,
                   "Value",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("A", 315), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

saveWorkbook(wb, file = "dashboard.xlsx", overwrite = TRUE)


# Enhanced Recovery

companies.to.show = c("Nutricia",
                      "Nestle",
                      "B. Braun Melsungen",
                      "Fresenius Kabi")

brands.to.show = c("Nutrison",
                   "Resource",
                   "Peptamen",
                   "Modulen",
                   "Isosource",
                   "Nutricomp",
                   "Fresubin")

segments.to.filter = 'PS == "Enhanced Recovery"'

result = dashTable(df,
                   "Volume",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

newOrder = c(4, 9, 3, 11, 10, 7, 6, 1, 8, 2, 5)

setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("O", 324), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

result = dashTable(df,
                   "Value",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("A", 324), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

saveWorkbook(wb, file = "dashboard.xlsx", overwrite = TRUE)

# Metabolics

companies.to.show = c("Nutricia")

brands.to.show = c("P-Am",
                   "Pku Nutri")

segments.to.filter = 'PS == "Metabolics"'

result = dashTable(df,
                   "Volume",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

newOrder = c(1, 3, 2)

setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("O", 340), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

result = dashTable(df,
                   "Value",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("A", 340), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

saveWorkbook(wb, file = "dashboard.xlsx", overwrite = TRUE)
