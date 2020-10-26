setwd("/home/serhii/Documents/Work/Nutricia/Data/202008")

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
Year = df[, max(Ynb, na.rm = TRUE)]
min.year = df[, min(Ynb, na.rm = TRUE)]
YTD.No = 8

df = df[(Ynb < Year) | (Ynb == Year & Mnb <= YTD.No)]

expected.periods = data.table(Ynb = c(
  sort(rep(min.year:(Year - 1), 12)), 
  rep(Year, YTD.No)),
Mnb = c(rep(1:12, (Year - min.year)), 
        1:YTD.No))

getTable = function(df, calculate.shares = NULL) {
    
    nc = length(df)
    
    df[, YTDLY := Reduce(`+`, .SD), .SDcols = (nc - YTD.No - 11):(nc - 12)]
    df[, YTD := Reduce(`+`, .SD), .SDcols = (nc - YTD.No + 1):nc]
    df[, MATLY := Reduce(`+`, .SD), .SDcols = (nc - 23):(nc - 12)]
    df[, MAT := Reduce(`+`, .SD), .SDcols = (nc - 11):nc]
    df[, L3MLY := Reduce(`+`, .SD), .SDcols = (nc - 14):(nc - 12)]
    df[, L3M := Reduce(`+`, .SD), .SDcols = (nc - 2):nc]
    df[, MCY := .SD, .SDcols = nc]
    
    df = df[, .SD, .SDcols = c(1, nc-12, nc-1, nc, (nc + 1):(nc + 7))]
    
    
    df = df[, paste0(names(df)[2:length(df)], "_MS") := lapply(.SD, function(x) x/sum(x)), 
            .SDcols = 2:length(df)]          
   
    
    df = df[, .SD, .SDcols = -c(2:(1 + (dim(df)[2]-1)/2))]
    
    df[, DeltaCMbps := 10000*Reduce(`-`, .SD), .SDcols = c(4, 3)]
    df[, DeltaYTDbps := 10000*Reduce(`-`, .SD), .SDcols = c(6, 5)]
    df[, DeltaMATbps := 10000*Reduce(`-`, .SD), .SDcols = c(8, 7)]
    df[, DeltaL3Mbps := 10000*Reduce(`-`, .SD), .SDcols = c(10, 9)]
    df[, DeltaCMLYbps := 10000*Reduce(`-`, .SD), .SDcols = c(4, 2)]
    
    df = setcolorder(df, c(1, 5:6, 13, 7:8, 14, 9:10, 15, 2, 4, 16, 3, 11, 12))
  
    return(df)
}

dashTable = function(df, measure, companies.to.show, 
                     brands.to.show, segments.to.filter = NULL) {

    df = df[eval(parse(text = segments.to.filter)), 
              .(Volume = sum(Volume), Value = sum(Value)),
              by = .(PS0, PS2, PS3, Company, Brand, Ynb, Mnb)]

    
    ## Check for the missing periods
    
    # df = data.table(A = c(1:3, 5,1:3, 5), B = c(1,1,1,1,2,2,2,2), C = "a")
    # expected = data.table(AE = 1:5, BE = c(rep(1, 5), rep(2,5)))
    # 
    df.aggregated = df[, .(Volume = sum(Volume), Value = sum(Value)),
                       by = .(PS0, Ynb, Mnb)]
    
    df.temp = df.aggregated[expected.periods, 
                 on = c("Ynb", "Mnb")][is.na(PS0)]
    
    df.temp[, `:=`(Volume = 0,
                   Value = 0,
                   PS0 = "Unknown",
                   PS2 = "Unknown",
                   PS3 = "Unknown",
                   Company = "Unknown",
                   Brand = "Unknown")] 
    
    df.temp = df.temp[, .SD, .SDcols = names(df)]
    df = rbindlist(list(df, df.temp))
    
    df = df[order(Ynb, Mnb)] 
    # df = rbindlist(list(df, df.temp))
    # 
    # df = df[order(B, A)]
    # df
    
    ##
    
        
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

### Segments

rows.order = c(
    "IMF + Dry Food + Puree",
    "IMF + Dry Food",
    "Foods (selected)",
    "IMF",
    "Base IF",
    "Base FO",
    "Base Gum",
    "Specials",
    "Digestive Comfort",
    "Hypoallergenic",
    "Allergy Treatment",
    "Preterm",
    "Goat",
    "Other TN",
    "Base Plus",
    "Dry Food",
    "Instant Cereals",
    "Cereal Biscuits",
    "Puree",
    "Meat Meal",
    "Veggie Meal",
    "Fruits",
    "Fruits Glass",
    "Fruits Pouch",
    "Fruits Tetra Pak",
    "AMN",
    "Recovery",
    "Enhanced Recovery",
    "Metabolics"
)

## Segments value

df1 = dcast.data.table(df, 
                       PS0 ~ Ynb + Mnb, 
                       fun = sum, 
                       value.var = "Value")

df2 = dcast.data.table(df[PS2 == "Dry Food"], 
                       PS2 ~ Ynb + Mnb, 
                       fun = sum, 
                       value.var = "Value")

df3 = dcast.data.table(df, 
                       PS3 ~ Ynb + Mnb, 
                       fun = sum, 
                       value.var = "Value")

df4 = dcast.data.table(df[PS0 != "AMN"], 
                       PS ~ Ynb + Mnb, 
                       fun = sum, 
                       value.var = "Value")

df5 = dcast.data.table(df[PS3 == "Fruits"], 
                       PS3 + Package ~ Ynb + Mnb, 
                       fun = sum, 
                       value.var = "Value")

df5[, PS3 := paste(PS3, Package)]
df5[, Package := NULL]

df6 = dcast.data.table(df[PS0 == "IMF" & PS3 == "Base"], 
                       PS3 + PS2 ~ Ynb + Mnb, 
                       fun = sum, 
                       value.var = "Value")

df6[, PS3 := paste(PS3, PS2)]
df6[, PS2 := NULL]

df7 = dcast.data.table(df[PS3 == "Instant Cereals" | PS3 == "Cereal Biscuits" |
                            PS3 == "Fruits" | PS3 == "Savoury Meal"], 
                       PS0 ~ Ynb + Mnb, 
                       fun = sum, 
                       value.var = "Value")

df7[, PS0 := "Foods (selected)"]

names(df1)[1] = "Segment"
names(df2)[1] = "Segment"    
names(df3)[1] = "Segment"    
names(df4)[1] = "Segment"    
names(df5)[1] = "Segment"    
names(df6)[1] = "Segment"   
names(df7)[1] = "Segment"    

df.segments = rbindlist(list(df1, df2, df3, df4, df5, df6, df7))

# part of the function
nc = length(df.segments)

df.segments[, YTDLY := Reduce(`+`, .SD), .SDcols = (nc - YTD.No - 11):(nc - 12)]
df.segments[, YTD := Reduce(`+`, .SD), .SDcols = (nc - YTD.No + 1):nc]
df.segments[, MATLY := Reduce(`+`, .SD), .SDcols = (nc - 23):(nc - 12)]
df.segments[, MAT := Reduce(`+`, .SD), .SDcols = (nc - 11):nc]
df.segments[, L3MLY := Reduce(`+`, .SD), .SDcols = (nc - 14):(nc - 12)]
df.segments[, L3M := Reduce(`+`, .SD), .SDcols = (nc - 2):nc]
df.segments[, MCY := .SD, .SDcols = nc]

df.segments = df.segments[, .SD, .SDcols = c(1, nc-12, nc-1, nc, (nc + 1):(nc + 7))]

## IMF+DF+Puree calculation
df1 = cbind(Segment = "IMF + Dry Food + Puree", 
            df.segments[Segment %in% c("IMF", "Dry Food", "Fruits", "Savoury Meal"), 
                        lapply(.SD, sum) , 
                        .SDcols = 2:dim(df.segments)[2]])
df.segments = rbindlist(list(df.segments, df1))


## IMF+DF calculation

df1 = cbind(Segment = "IMF + Dry Food", 
            df.segments[Segment %in% c("IMF", "Dry Food"), 
                        lapply(.SD, sum) , 
                        .SDcols = 2:dim(df.segments)[2]])
df.segments = rbindlist(list(df.segments, df1))



## Other TN calculation

df1 = cbind(Segment = "Other TN", 
            df.segments[Segment %in% c("DR-NL", "Anti Reflux", "Soy"), 
                        lapply(.SD, sum) , 
                        .SDcols = 2:dim(df.segments)[2]])
df.segments = rbindlist(list(df.segments, df1))

## Puree

df1 = cbind(Segment = "Puree", 
            df.segments[Segment %in% c("Fruits", "Savoury Meal"), 
                        lapply(.SD, sum) , 
                        .SDcols = 2:dim(df.segments)[2]])
df.segments = rbindlist(list(df.segments, df1))

## Extra columns
df.segments[, CMEvo := Reduce(`/`, .SD) - 1, .SDcols = c(4, 3)]
df.segments[, YTDEvo := Reduce(`/`, .SD) - 1, .SDcols = c(6, 5)]
df.segments[, MATEvo := Reduce(`/`, .SD) - 1, .SDcols = c(8, 7)]
df.segments[, L3MEvo := Reduce(`/`, .SD) - 1, .SDcols = c(10, 9)]
df.segments[, CMLYEvo := Reduce(`/`, .SD) - 1, .SDcols = c(4, 2)]

cols.excl = c("Segment", "YTDEvo", "MATEvo", "L3MEvo", "CMLYEvo", "CMEvo")
cols = names(df.segments)[!(names(df.segments) %in% cols.excl)]

df.segments[, (cols) := lapply(.SD, function(x) x/1000), .SDcols = cols]

result = setcolorder(df.segments, c(1, 5:6, 13, 7:8, 14, 9:10, 15, 2, 4, 16, 3, 11, 12))

result = df.segments[Segment %in% rows.order]

result = result[match(rows.order, Segment),]

writeData(wb, 
          "Sheet1", 
          "Segments", 
          rowNames = FALSE, 
          xy = c("A", 2))

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("A", 4), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight1")

# Segments Volume

df1 = dcast.data.table(df, 
                 PS0 ~ Ynb + Mnb, 
                 fun = sum, 
                 value.var = "Volume")

df2 = dcast.data.table(df[PS2 == "Dry Food"], 
                       PS2 ~ Ynb + Mnb, 
                       fun = sum, 
                       value.var = "Volume")

df3 = dcast.data.table(df, 
                       PS3 ~ Ynb + Mnb, 
                       fun = sum, 
                       value.var = "Volume")

df4 = dcast.data.table(df[PS0 != "AMN"], 
                       PS ~ Ynb + Mnb, 
                       fun = sum, 
                       value.var = "Volume")

df5 = dcast.data.table(df[PS3 == "Fruits"], 
                       PS3 + Package ~ Ynb + Mnb, 
                       fun = sum, 
                       value.var = "Volume")

df5[, PS3 := paste(PS3, Package)]
df5[, Package := NULL]

df6 = dcast.data.table(df[PS0 == "IMF" & PS3 == "Base"], 
                        PS3 + PS2 ~ Ynb + Mnb, 
                        fun = sum, 
                        value.var = "Volume")

df6[, PS3 := paste(PS3, PS2)]
df6[, PS2 := NULL]

df7 = dcast.data.table(df[PS3 == "Instant Cereals" | PS3 == "Cereal Biscuits" |
                            PS3 == "Fruits" | PS3 == "Savoury Meal"], 
                       PS0 ~ Ynb + Mnb, 
                       fun = sum, 
                       value.var = "Volume")

df7[, PS0 := "Foods (selected)"]

names(df1)[1] = "Segment"
names(df2)[1] = "Segment"    
names(df3)[1] = "Segment"    
names(df4)[1] = "Segment"    
names(df5)[1] = "Segment"    
names(df6)[1] = "Segment"    
names(df7)[1] = "Segment"    

df.segments = rbindlist(list(df1, df2, df3, df4, df5, df6, df7))

# part of the function
nc = length(df.segments)

df.segments[, YTDLY := Reduce(`+`, .SD), .SDcols = (nc - YTD.No - 11):(nc - 12)]
df.segments[, YTD := Reduce(`+`, .SD), .SDcols = (nc - YTD.No + 1):nc]
df.segments[, MATLY := Reduce(`+`, .SD), .SDcols = (nc - 23):(nc - 12)]
df.segments[, MAT := Reduce(`+`, .SD), .SDcols = (nc - 11):nc]
df.segments[, L3MLY := Reduce(`+`, .SD), .SDcols = (nc - 14):(nc - 12)]
df.segments[, L3M := Reduce(`+`, .SD), .SDcols = (nc - 2):nc]
df.segments[, MCY := .SD, .SDcols = nc]

df.segments = df.segments[, .SD, .SDcols = c(1, nc-12, nc-1, nc, (nc + 1):(nc + 7))]

## IMF+DF+Puree calculation
df1 = cbind(Segment = "IMF + Dry Food + Puree", 
            df.segments[Segment %in% c("IMF", "Dry Food", "Fruits", "Savoury Meal"), 
                        lapply(.SD, sum) , 
                        .SDcols = 2:dim(df.segments)[2]])
df.segments = rbindlist(list(df.segments, df1))


## IMF+DF calculation

df1 = cbind(Segment = "IMF + Dry Food", 
            df.segments[Segment %in% c("IMF", "Dry Food"), 
                        lapply(.SD, sum) , 
                        .SDcols = 2:dim(df.segments)[2]])
df.segments = rbindlist(list(df.segments, df1))

## Other TN calculation

df1 = cbind(Segment = "Other TN", 
            df.segments[Segment %in% c("DR-NL", "Anti Reflux", "Soy"), 
                        lapply(.SD, sum) , 
                        .SDcols = 2:dim(df.segments)[2]])
df.segments = rbindlist(list(df.segments, df1))

## Puree

df1 = cbind(Segment = "Puree", 
            df.segments[Segment %in% c("Fruits", "Savoury Meal"), 
                        lapply(.SD, sum) , 
                        .SDcols = 2:dim(df.segments)[2]])
df.segments = rbindlist(list(df.segments, df1))

## Extra columns
df.segments[, CMEvo := Reduce(`/`, .SD) - 1, .SDcols = c(4, 3)]
df.segments[, YTDEvo := Reduce(`/`, .SD) - 1, .SDcols = c(6, 5)]
df.segments[, MATEvo := Reduce(`/`, .SD) - 1, .SDcols = c(8, 7)]
df.segments[, L3MEvo := Reduce(`/`, .SD) - 1, .SDcols = c(10, 9)]
df.segments[, CMLYEvo := Reduce(`/`, .SD) - 1, .SDcols = c(4, 2)]

result = setcolorder(df.segments, c(1, 5:6, 13, 7:8, 14, 9:10, 15, 2, 4, 16, 3, 11, 12))

result = df.segments[Segment %in% rows.order]

result = result[match(rows.order, Segment),]

writeData(wb, 
          "Sheet1", 
          "Segments", 
          rowNames = FALSE, 
          xy = c("R", 2))

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("R", 4), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight1")

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
          xy = c("R", 35))

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("R", 37), 
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

writeData(wb, 
          "Sheet1", 
          "IMF + Dry Food", 
          rowNames = FALSE, 
          xy = c("R", 60))

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("R", 62), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

result = dashTable(df,
                   "Value",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

result = add.nutricia.wo.malysh(result, TRUE)

setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

writeData(wb, 
          "Sheet1", 
          "IMF + Dry Food", 
          rowNames = FALSE, 
          xy = c("A", 60))

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("A", 62), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")


# Foods (selected)
companies.to.show = c(
  "Nutricia",
  "Nestle",
  "Vitmark",
  "Hipp",
  "ADP",
  "Droga Kolinska",
  "Hame",
  "Ekoniya",
  "Khorolskii Mk"
)

brands.to.show = c("Nutrilon",
                   "Milupa",
                   "Gerber")
segments.to.filter = 'PS3 == "Instant Cereals" | PS3 == "Cereal Biscuits" |
PS3 == "Fruits" | PS3 == "Savoury Meal"'


result = dashTable(df, "Volume", companies.to.show, brands.to.show, segments.to.filter)
newOrder = c(8, 9, 11, 10, 7, 12, 4, 1, 6, 5, 3, 2)
setorder(result[, .r := newOrder], .r)[, .r := NULL]

writeData(wb, 
          "Sheet1", 
          "Foods (selected)", 
          rowNames = FALSE, 
          xy = c("R", 370))

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("R", 373), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")


result = dashTable(df, "Value", companies.to.show, brands.to.show, segments.to.filter)

setorder(result[, .r := newOrder], .r)[, .r := NULL]

writeData(wb, 
          "Sheet1", 
          "Foods (selected)", 
          rowNames = FALSE, 
          xy = c("A", 370))

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("A", 373), 
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

writeData(wb, 
          "Sheet1", 
          "IMF", 
          rowNames = FALSE, 
          xy = c("R", 90))

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("R", 92), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")


result = dashTable(df,
                   "Value",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

setorder(result[, .r := newOrder], .r)[, .r := NULL]

writeData(wb, 
          "Sheet1", 
          "IMF", 
          rowNames = FALSE, 
          xy = c("A", 90))

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

writeData(wb, 
          "Sheet1", 
          "IF Base", 
          rowNames = FALSE, 
          xy = c("R", 109))

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("R", 111), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

result = dashTable(df,
                   "Value",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

setorder(result[, .r := newOrder], .r)[, .r := NULL]

writeData(wb, 
          "Sheet1", 
          "IF Base", 
          rowNames = FALSE, 
          xy = c("A", 109))

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

writeData(wb, 
          "Sheet1", 
          "FO Base", 
          rowNames = FALSE, 
          xy = c("R", 127))

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("R", 129), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")


result = dashTable(df, "Value", companies.to.show, brands.to.show, segments.to.filter)

setorder(result[, .r := newOrder], .r)[, .r := NULL]

writeData(wb, 
          "Sheet1", 
          "FO Base", 
          rowNames = FALSE, 
          xy = c("A", 127))

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

writeData(wb, 
          "Sheet1", 
          "Gum Base", 
          rowNames = FALSE, 
          xy = c("R", 145))

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("R", 147), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

result = dashTable(df, "Value", companies.to.show, brands.to.show, segments.to.filter)

setorder(result[, .r := newOrder], .r)[, .r := NULL]

writeData(wb, 
          "Sheet1", 
          "Gum Base", 
          rowNames = FALSE, 
          xy = c("A", 145))

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

writeData(wb, 
          "Sheet1", 
          "Specials", 
          rowNames = FALSE, 
          xy = c("R", 162))

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("R", 164), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

result = dashTable(df,
                   "Value",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

setorder(result[, .r := newOrder], .r)[, .r := NULL]

writeData(wb, 
          "Sheet1", 
          "Specials", 
          rowNames = FALSE, 
          xy = c("A", 162))

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

writeData(wb, 
          "Sheet1", 
          "Dry Food", 
          rowNames = FALSE, 
          xy = c("R", 176))

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("R", 178), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

result = dashTable(df,
                   "Value",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

setorder(result[, .r := newOrder], .r)[, .r := NULL]

writeData(wb, 
          "Sheet1", 
          "Dry Food", 
          rowNames = FALSE, 
          xy = c("A", 176))

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

writeData(wb, 
          "Sheet1", 
          "Cereals", 
          rowNames = FALSE, 
          xy = c("R", 193))

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("R", 195), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")


result = dashTable(df, "Value", companies.to.show, brands.to.show, segments.to.filter)

setorder(result[, .r := newOrder], .r)[, .r := NULL]

writeData(wb, 
          "Sheet1", 
          "Cereals", 
          rowNames = FALSE, 
          xy = c("A", 193))

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

writeData(wb, 
          "Sheet1", 
          "Biscuits", 
          rowNames = FALSE, 
          xy = c("R", 210))

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("R", 212), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")


result = dashTable(df, "Value", companies.to.show, brands.to.show, segments.to.filter)

setorder(result[, .r := newOrder], .r)[, .r := NULL]

writeData(wb, 
          "Sheet1", 
          "Biscuits", 
          rowNames = FALSE, 
          xy = c("A", 210))

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("A", 212), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

# Puree

# companies.to.show = c(
#     "Vitmark",
#     "Nestle",
#     "ADP",
#     "Hipp",
#     "Hame",
#     "Nutricia",
#     "Ekoniya",
#     "Profitzernomarket",
#     "Belfood"
# )

companies.to.show = c("Hame")

brands.to.show = c("Bebivita",
"Chudo-Chado",
"Gerber",
"Hipp",
"Karapuz",
"Lozhka V Ladoshke",
"Malenkoe Schaste",
"Malyatko",
"Milupa")


segments.to.filter = 'PS3 == "Fruits" | PS3 == "Savoury Meal"'
    
result = dashTable(df,
                   "Volume",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

# newOrder = c(7, 6, 9, 5, 4, 1, 8, 3, 2) # for companies
newOrder = c(10, 3, 4, 6, 2, 5, 9, 7, 8, 2) # for brands


setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

writeData(wb, 
          "Sheet1", 
          "Puree", 
          rowNames = FALSE, 
          xy = c("R", 222))

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("R", 224), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

result = dashTable(df,
                   "Value",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

writeData(wb, 
          "Sheet1", 
          "Puree", 
          rowNames = FALSE, 
          xy = c("A", 222))

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("A", 224), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

# saveWorkbook(wb, file = "dashboard.xlsx", overwrite = TRUE)

# Meat Meal

# companies.to.show = c("Hame",
#                       "Hipp",
#                       "Nutricia",
#                       "Nestle",
#                       "Profitzernomarket",
#                       "ADP")

companies.to.show = c("")

brands.to.show = c(
  "Bebivita",
  "Fruta Pyureshka",
  "Gerber",
  "Hamanek",
  "Hame",
  "Hipp",
  "Karapuz",
  "Lozhka V Ladoshke",
  "Milupa"
)

segments.to.filter = 'PS == "Meat Meal"'

result = dashTable(df,
                   "Volume",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

# newOrder = c(5, 2,3, 4, 6, 1) for companies

newOrder = c(9, 3, 6, 5, 4, 8, 7, 2, 1)

setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

writeData(wb, 
          "Sheet1", 
          "Meat Meal", 
          rowNames = FALSE, 
          xy = c("R", 236))

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("R", 238), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

result = dashTable(df,
                   "Value",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

writeData(wb, 
          "Sheet1", 
          "Meat Meal", 
          rowNames = FALSE, 
          xy = c("A", 236))

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("A", 238), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

# saveWorkbook(wb, file = "dashboard.xlsx", overwrite = TRUE)

# Veggie Meal

# companies.to.show = c(
#     "Nestle",
#     "Hipp",
#     "Nutricia",
#     "Profitzernomarket",
#     "Belfood",
#     "Oasis Group",
#     "ADP",
#     "Ekoniya"
# )

companies.to.show = c("")
brands.to.show = c(
  "Bambolina",
  "Bebivita",
  "Gerber",
  "Hipp",
  # "Karapuz",
  "Lozhka V Ladoshke",
  "Malenkoe Schaste",
  "Malyatko",
  "Milupa"
)
segments.to.filter = 'PS == "Veggie Meal"'

result = dashTable(df,
                   "Volume",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

# newOrder = c(6, 5, 4, 8, 2, 7, 3, 1)
newOrder = c(8, 3, 4, 5, 6, 1, 7, 2)

setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

writeData(wb, 
          "Sheet1", 
          "Veggie Meal", 
          rowNames = FALSE, 
          xy = c("R", 250))

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("R", 252), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

result = dashTable(df,
                   "Value",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

writeData(wb, 
          "Sheet1", 
          "Veggie Meal", 
          rowNames = FALSE, 
          xy = c("A", 250))

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("A", 252), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

# saveWorkbook(wb, file = "dashboard.xlsx", overwrite = TRUE)

# Fruit Jar/Glass

# companies.to.show = c(
#     "Vitmark",
#     "ADP",
#     "Nestle",
#     "Ekoniya",
#     "Nutricia",
#     "Hipp",
#     "Hame",
#     "Belfood",
#     "Oasis Group"
# )

companies.to.show = c("")
brands.to.show = c(
  "Bambolina",
  "Chudo-Chado",
  "Gerber",
  "Hame",
  "Hipp",
  "Karapuz",
  "Malenkoe Schaste",
  "Malyatko",
  "Milupa"
)
segments.to.filter = 'PS3 == "Fruits" & Package == "Glass"'

result = dashTable(df,
                   "Volume",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

# newOrder = c(7, 9, 1, 6, 3, 4, 5, 2, 8)
newOrder = c(9, 3, 2, 6, 8, 5, 4, 7, 1)

setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

writeData(wb, 
          "Sheet1", 
          "Fruit Jar/Glass", 
          rowNames = FALSE, 
          xy = c("R", 264))

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("R", 266), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

result = dashTable(df,
                   "Value",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

writeData(wb, 
          "Sheet1", 
          "Fruit Jar/Glass", 
          rowNames = FALSE, 
          xy = c("A", 264))

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("A", 266), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

# saveWorkbook(wb, file = "dashboard.xlsx", overwrite = TRUE)

# Fruit Pouch

# companies.to.show = c("Vitmark",
#                       "Nestle",
#                       "Hipp",
#                       "Hame",
#                       "Nutricia",
#                       "Profitzernomarket",
#                       "Belfood")

companies.to.show = c("")

brands.to.show = c("Chudo-Chado",
                   "Gerber",
                   "Hamanek",
                   "Hipp",
                   "Lozhka V Ladoshke",
                   "Milupa")

segments.to.filter = 'PS3 == "Fruits" & Package == "Pouch"'

result = dashTable(df,
                   "Volume",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

# newOrder = c(5, 7, 4, 2, 3, 6, 1)
newOrder = c(6, 2, 1, 4, 3, 5)

setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

writeData(wb, 
          "Sheet1", 
          "Fruit Pouch", 
          rowNames = FALSE, 
          xy = c("R", 278))

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("R", 280), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

result = dashTable(df,
                   "Value",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

writeData(wb, 
          "Sheet1", 
          "Fruit Pouch", 
          rowNames = FALSE, 
          xy = c("A", 278))

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("A", 280), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

# saveWorkbook(wb, file = "dashboard.xlsx", overwrite = TRUE)

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

writeData(wb, 
          "Sheet1", 
          "AMN", 
          rowNames = FALSE, 
          xy = c("R", 292))

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("R", 294), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

result = dashTable(df,
                   "Value",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

writeData(wb, 
          "Sheet1", 
          "AMN", 
          rowNames = FALSE, 
          xy = c("A", 292))

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("A", 294), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

# saveWorkbook(wb, file = "dashboard.xlsx", overwrite = TRUE)

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

writeData(wb, 
          "Sheet1", 
          "Recovery", 
          rowNames = FALSE, 
          xy = c("R", 313))

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("R", 315), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

result = dashTable(df,
                   "Value",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

writeData(wb, 
          "Sheet1", 
          "Recovery", 
          rowNames = FALSE, 
          xy = c("A", 313))

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("A", 315), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

# saveWorkbook(wb, file = "dashboard.xlsx", overwrite = TRUE)


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

writeData(wb, 
          "Sheet1", 
          "Enhanced Recovery", 
          rowNames = FALSE, 
          xy = c("R", 322))

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("R", 324), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

result = dashTable(df,
                   "Value",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

writeData(wb, 
          "Sheet1", 
          "Enhanced Recovery", 
          rowNames = FALSE, 
          xy = c("A", 322))

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("A", 324), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

# saveWorkbook(wb, file = "dashboard.xlsx", overwrite = TRUE)

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

writeData(wb, 
          "Sheet1", 
          "Metabolics", 
          rowNames = FALSE, 
          xy = c("R", 338))

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("R", 340), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

result = dashTable(df,
                   "Value",
                   companies.to.show,
                   brands.to.show,
                   segments.to.filter)

setorder(result[, .r := order(newOrder)], .r)[, .r := NULL]

writeData(wb, 
          "Sheet1", 
          "Metabolics", 
          rowNames = FALSE, 
          xy = c("A", 338))

writeDataTable(wb, 
               "Sheet1", 
               x = result, 
               xy = c("A", 340), 
               rowNames = FALSE,
               tableStyle = "TableStyleLight9")

saveWorkbook(wb, file = "dashboard.xlsx", overwrite = TRUE)






