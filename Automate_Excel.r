## -------------------------------------------
## read in data, reformat, write to excel
## 
## -------------------------------------------

library(RDCOMClient)
library(plyr)

options(stringsAsFactors = FALSE)

## define study and report type
depos <- ''

## read in RDCOMClient excel functions
source("excelUtils.R")

path <- paste("yourFilePath", depos, "subFolder", sep='')
setwd(path)


## function to convert column index to letter for excel
letter_map <- as.data.frame(cbind(toupper(letters), 1:26))
names(letter_map) <- c('letter', 'number')
excel_er_ate <- function(number)
	{
	modu <- number %/% 26
	rem  <- number %% 26

	if(rem !=0 )
		{
		letter <- subset(letter_map, number==rem)$letter
		} else
		{
		letter <- 'Z'	## ex:  number=52
		}

	if(modu != 0)
		{
		if(rem != 0)
			{
			letter_prefix <- subset(letter_map, number==modu)$letter
			} else
			{
			letter_prefix <- subset(letter_map, number==modu-1)$letter
			}

		prefix <- rep(letter_prefix, 1)
		letter <- paste(letter_prefix, letter, sep='')
		}

	return(letter)
	}


## determine how many reports to loop through
reports <- NULL

# codes to create a list of report names

reports <- as.data.frame(reports)
names(reports) <- c('table_id', 'report_name')

for(r in 1:nrow(reports))
	{
	table_id <- reports$table_id[r]
	report_name <- reports$report_name[r]

	setwd(path)
	data_file    <- list.files()[grep(paste(table_id, "data.csv", sep=''), list.files())]
	outlier_file <- list.files()[grep(paste(table_id, "outliers.csv", sep=''), list.files())]

	data3 <- read.csv(data_file)
	data3[is.na(data3)] <- ''

	outliers <- read.csv(outlier_file)
	outliers[is.na(outliers)] <- ''

	## add blank columns so we can have ONE version of this file for any input dataset
	common_fields <- c('list of column names of your file')

	missing_cols <- common_fields[!common_fields %in% names(data3)]
	start_names <- names(data3)
	temp_data2 <- data3
	if(length(missing_cols) > 0)
		{
		for(i in 1:length(missing_cols))
			{
			temp_data2 <- cbind(temp_data2, rep('', nrow(temp_data2)))
			}

		names(temp_data2) <- c(start_names, missing_cols)
		} else
		{
		temp_data2 <- data3
		}

	data3 <- temp_data2


	if('endpoint' %in% names(data3))
		{
		pointflag <- 'yes'

		## transpose data, add flag for outliers
		data3_list <- dlply(data3, .(endpoint))

		thin_list <- list()
		for(j in 1:length(data3_list))
			{
			temp_list <- data3_list[[j]]
			endpoint <- temp_list$endpoint[1]
			endpoint <- gsub(' ', '_', endpoint)
			names(temp_list)[names(temp_list) == 'numeric_value'] <- endpoint
			temp_list <- temp_list[,c('list of column names of your file')]

			thin_list[[j]] <- temp_list
			}

		## collapse list, merge enpdoints horizontally
		## use first index as template
		out <- thin_list[[1]]
		if(length(thin_list) > 1)
			{
			for(k in 2:length(thin_list))
				{
				temp_list <- thin_list[[k]]

				## merge by everything except endpoint
				out <- merge(out, temp_list, by=names(out)[1:18], all=TRUE)
				}
			}
		} else
		{
		pointflag <- 'no'
		out <- data3[,c('list of column names of your file')]
		}


	out$outlier_present <- 'F'


	## loop through sex, create excel sheet
	xls <- COMCreate("Excel.Application")
	xls[["Visible"]] <- TRUE
	wb <- xls[["Workbooks"]]$Add(1)

	for(s in unique(out$sex))
		{
		data_temp <- subset(out, sex==s)
		rownames(data_temp) <- 1:nrow(data_temp)
		data_temp[is.na(data_temp)] <- ''

		out_temp <- subset(outliers, sex==s)
 

		## create excel
		sh <- wb[["Worksheets"]]$Add()
		tab_name <- paste(s, report_name, sep='_')
		if(nchar(tab_name) > 31) { tab_name <- substr(tab_name, 1, 31) } 	## 31 char limit on excel tab names
		sh[["Name"]] <- tab_name

		## write data
		exportDataFrame(data_temp, at=sh$Range("A1"))

		## find outliers, highlight in red
		if(nrow(out_temp) > 0)
			{
			for(ot in 1:nrow(out_temp))
				{
				sex <- s
				gen <- out_temp$generation[ot]
				sel <- out_temp$selection[ot]
				lit <- out_temp$litter_name[ot]
				phase <- out_temp$phase_type[ot]
				time <- out_temp$phase_time[ot]
				start <- out_temp$phase_start[ot]
				end <- out_temp$phase_end[ot]

				dose_j <- out_temp$dose[ot]
				value <- out_temp$numeric_value[ot]
				if(pointflag == 'yes')
					{
					point <- out_temp$endpoint[ot]
					}

				## get cell location
				if(pointflag == 'yes')
					{
					point <- gsub(' ', '_', point)
					} else
					{
					point <- 'numeric_value'
					}

				col_index <- which(names(data_temp) == point)
				data_temp[,col_index] <- as.numeric(data_temp[,col_index])
				row_index <- as.numeric(rownames(subset(data_temp, dose==dose_j & selection==sel & generation==gen & litter_name==lit & phase_type==phase & phase_time==time & phase_start==start & phase_end==end)[value == subset(data_temp, dose==dose_j & selection==sel & generation==gen & litter_name==lit & phase_type==phase & phase_time==time & phase_start==start & phase_end==end)[,col_index],])) + 1  ## add one for headers in excel
				row_index <- row_index[!is.na(row_index)]

				## excel-er-ate col_index into letter
				col_letter <- excel_er_ate(col_index)
				cell <- paste(col_letter, row_index, sep='')

				## highlight outliers!  May have multiple rows with same value (multiple outliers), loop through
				for(z in 1:length(cell))
					{
					temp_cell <- cell[z]
					highlight <- sh$Range(paste(temp_cell, ":", temp_cell, sep=''))
					highlight2 <- highlight$Font()
					highlight2[["Bold"]] <- TRUE
					highlight2[["Color"]] <- "255" 	## red

					## set outlier flag to TRUE
					flag_index <- grep('outlier_present', names(data_temp))
					flag_letter <- excel_er_ate(flag_index)
					temp_cell <- paste(flag_letter, row_index[z], sep='')

					flag <- sh$Range(paste(temp_cell, ":", temp_cell, sep=''))
					flag[["Value"]] <- 'T'
					}
				}
			}
		}
	## clean up default sheet
	xls$Sheets("Sheet1")$Delete()

	## create folder for outliers if it doesn't exist
	if(!'formatted_outlier_files' %in% dir())
	{
	  dir.create('formatted_outlier_files')
	}
	setwd('formatted_outlier_files')


filename <- paste(getwd(), "/", depos, "_", report_name, "_outliers.xlsx", sep='')

wb$SaveAs(filename)
xls$Quit()		## close excel automatically
}
