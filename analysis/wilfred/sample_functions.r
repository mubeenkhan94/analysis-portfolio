# Samples from Wilfred

#' Return the fiscal year for a date as a 4-digit integer
#' 
#' @param x A date in "YYYY-MM-DD" format (default is system date at function runtime)
#' @return A 4-digit integer
#' @examples
#' get_fiscal_year()
#' get_fiscal_year("2001-09-24")
get_fiscal_year <- function(date = Sys.Date()) {
    fiscal_year <- as.integer(as.yearmon(date) - 8/12 + 1)
    return(fiscal_year)
}

#' Return a list of the month-years for a fiscal year
#' 
#' @param x An integer, leave blank for the calendar year at function runtime
#' @param y A month cutoff for the list in %b %y format, leave blank for calendar month at runtime
#' @return A list of months
#' @examples
#' get_months_of_fiscal_years_to_date()
#' get_months_of_fiscal_years_to_date(2021)
#' get_months_of_fiscal_years_to_date(2021, "Mar 21")
get_months_of_fiscal_years_to_date <- function(
    fiscal_year = format(Sys.Date(), "%Y"), 
    end_month = format(Sys.Date(), "%b %y")
    ) {
    starting_month_date <- as.Date(paste0(as.numeric(fiscal_year)-1, "-09-01"))
    end_month_date <- as.Date(parse_date_time(end_month, "%b %y"))
    if (end_month_date >= as.Date(paste0(fiscal_year,"-09-01"))) {
        end_month_date <- as.Date(paste0(fiscal_year,"-08-01"))
    } else {}
    month_list <- format(seq(starting_month_date, end_month_date, by = "month"), "%b %y")
    return(month_list)
}

#' Removes headers and footers from Excel exports from CUIC
#' 
#' @param x A dataframe
#' @return A dataframe without the CUIC header and footer
clean_report_from_cuic <- function(report_from_cuic) {
    if (grepl("Generated", report_from_cuic[nrow(report_from_cuic)-1,])[1] == TRUE) {
        footer_row_numbers <- c(
            nrow(report_from_cuic),
            nrow(report_from_cuic) - 1,
            nrow(report_from_cuic) -2
        )
        report_from_cuic_no_footer <-
            report_from_cuic[-footer_row_numbers,]
        return(report_from_cuic_no_footer)
    } else {
        return(report_from_cuic)
    }
}

#' Takes an ABA report from CUIC, removes headers and footers, renames columns, converts formats where necessary, and returns a dataframe
#' 
#' @param x A dataframe
#' @return A dataframe without the CUIC header and footer, with columns renamed, formats converted
clean_aba_report_from_cuic <- function(aba_report_from_cuic, new_date_type = FALSE) {
    if (class(aba_report_from_cuic) == "data.frame") {
        report_from_cuic_clean <- 
            clean_report_from_cuic(aba_report_from_cuic)
    } else if (class(aba_report_from_cuic) == "list") {
        report_from_cuic_clean <-
            lapply(aba_report_from_cuic, 
                    clean_report_from_cuic) %>% 
                bind_rows
    }
    if (new_date_type == TRUE) {
        report_from_cuic_clean <- report_from_cuic_clean %>%
            mutate(date_time = parse_new_cuic_date_format(DateTime), date = parse_new_cuic_date_format(DATE, "date"))
    } else {
        report_from_cuic_clean <- report_from_cuic_clean %>%
            mutate(date_time = convertToDateTime(DateTime), date = convertToDate(DATE))
    }
    report_from_cuic_clean <-
        report_from_cuic_clean %>%
            mutate(precision_queue = tolower(Precision.Queue),
                    # date = date_time %>% format("%Y-%m-%d") %>% as.Date,
                    weekday = weekdays(date),
                    time = date_time %>% format("%H:%M:%S") %>% hms %>% as_datetime,
                    hour = hour(time),
                    service_level = Service.Level,
                    hold_time = convertToDateTime(Hold.Time, origin = "1970-01-01"),
                    speed_of_answer_mean = convertToDateTime(Avg.Speed.of.Answer, origin = "1970-01-01"),
                    calls_offered = Total,
                    calls_handled = Handled,
                    calls_abandoned = Aban,
                    calls_answered = Answered,
                    handle_time_mean = convertToDateTime(Avg.Handle.Time, origin = "1970-01-01"),
                    longest_queued = convertToDateTime(Longest.Queued, origin = "1970-01-01"),
                    max_queued = Max.Queued,
                    service_level_abandoned = Service.Level.Abandon
                    ) %>%
            select(precision_queue,
                    date,
                    weekday,
                    date_time,
                    time,
                    hour,
                    service_level,
                    hold_time,
                    speed_of_answer_mean,
                    calls_offered,
                    calls_handled,
                    calls_abandoned,
                    calls_answered,
                    handle_time_mean,
                    longest_queued,
                    max_queued,
                    service_level_abandoned)
    queue_names <-
        c("clsma",
          "corp",
          "derm",
          "gim",
          "intmed",
          "cim",
          "ped",
          "student",
          "travel")
    #### rename queue names to readable English
    for (queue_name in queue_names) {
        report_from_cuic_clean$precision_queue[grepl(queue_name, report_from_cuic_clean$precision_queue)] <- toupper(queue_name)
    }
    return(report_from_cuic_clean)
}

#' Takes all of the aba reports saved for a month in the working directory and returns a single dataframe
#' 
#' @param month A character string of the month in question
#' @return A dataframe of all of the aba reports for a month
get_list_of_aba_reports_for_a_month <- function(month) {
    #### adapt month entry to filename convention
    month <- tolower(month)
    month <- sub("(\\w+) (\\w+)", "\\1_\\2", month)
    #### get matching files
    files <- list.files()[grepl("aba_report", list.files()) & grepl(month, list.files())]
    #### create the list of dataframes
    read_filename_from_directory <- function(filename) {
        result <- tryCatch((read.xlsx(filename)), error = function(e) {})
        return(result)
    }
    get_aba_file_for_queue <- function(filename) {
        result <- read_filename_from_directory(paste0(filename))
        return(result)
    }
    aba_report_list <- lapply(files, get_aba_file_for_queue)
    aba_report_list <- Filter(Negate(is.null), aba_report_list)
    return(aba_report_list)
}
 
#' Helper function that generates an ABA summary similar to the end of day report
#' 
#' @param x A dataframe from CUIC, cleaned
#' @return A dataframe similar to the end of day report
generate_aba_summary_standard <- function(report_from_cuic_clean) {
    aba_summary_standard_by_queue <-
        report_from_cuic_clean %>%
            group_by(precision_queue) %>%
            summarise(
                service_level_mean = mean(na.omit(service_level)),
                calls_offered_total = sum(na.omit(calls_offered)),
                calls_answered_total = sum(na.omit(calls_answered)),
                calls_abandoned_total = sum(na.omit(calls_abandoned)),
                handle_time_mean = mean(na.omit(handle_time_mean)),
                speed_of_answer_mean = mean(na.omit(speed_of_answer_mean)),
                hold_time_mean = mean(na.omit(hold_time))
            ) %>%
            ungroup() %>%
            mutate(aba = calls_abandoned_total/calls_offered_total)
    aba_summary_standard_overall <-
        report_from_cuic_clean %>%
            summarise(
                service_level_mean = mean(na.omit(service_level)),
                calls_offered_total = sum(na.omit(calls_offered)),
                calls_answered_total = sum(na.omit(calls_answered)),
                calls_abandoned_total = sum(na.omit(calls_abandoned)),
                handle_time_mean = mean(na.omit(handle_time_mean)),
                speed_of_answer_mean = mean(na.omit(speed_of_answer_mean)),
                hold_time_mean = mean(na.omit(hold_time))
            ) %>%
            mutate(aba = calls_abandoned_total/calls_offered_total,
                    precision_queue = "total_overall")
    aba_summary_standard_by_queue <- 
        aba_summary_standard_by_queue[c(1,2,9,3,4,5,6,7,8)]
    aba_summary_standard_overall <- 
        aba_summary_standard_overall[c(9,1,8,2,3,4,5,6,7)]
    aba_summary_standard <-
        rbind(aba_summary_standard_by_queue,
                aba_summary_standard_overall)
    return(aba_summary_standard)
}

#' Helper function that generates an ABA report summarized by queue and date, including total/overall
#' 
#' @param x A dataframe from CUIC, cleaned
#' @return A dataframe summarizing by queue and date, including total/overall
generate_aba_summary_by_date <- function(report_from_cuic_clean) {
    aba_summary_by_date_by_queue <-
        report_from_cuic_clean %>%
            group_by(precision_queue, date) %>%
            summarise(
                service_level_mean = mean(na.omit(service_level)),
                calls_offered_total = sum(na.omit(calls_offered)),
                calls_answered_total = sum(na.omit(calls_answered)),
                calls_abandoned_total = sum(na.omit(calls_abandoned)),
                handle_time_mean = mean(na.omit(handle_time_mean)),
                speed_of_answer_mean = mean(na.omit(speed_of_answer_mean)),
                hold_time_mean = mean(na.omit(hold_time))
            ) %>%
            ungroup() %>%
            mutate(aba = calls_abandoned_total/calls_offered_total)
    aba_summary_by_date_overall <-
        report_from_cuic_clean %>%
            group_by(date) %>%
            summarise(
                service_level_mean = mean(na.omit(service_level)),
                calls_offered_total = sum(na.omit(calls_offered)),
                calls_answered_total = sum(na.omit(calls_answered)),
                calls_abandoned_total = sum(na.omit(calls_abandoned)),
                handle_time_mean = mean(na.omit(handle_time_mean)),
                speed_of_answer_mean = mean(na.omit(speed_of_answer_mean)),
                hold_time_mean = mean(na.omit(hold_time))
            ) %>%
            ungroup() %>%
            mutate(aba = calls_abandoned_total/calls_offered_total,
                    precision_queue = "total_overall")
    aba_summary_by_date_by_queue <- 
        aba_summary_by_date_by_queue[c(1,2,10,3,4,5,6,7,8,9)]
    aba_summary_by_date_overall <- 
        aba_summary_by_date_overall[c(10,1,2,9,3,4,5,6,7,8)]
    aba_summary_by_date <-
        rbind(aba_summary_by_date_by_queue,
                aba_summary_by_date_overall) %>%
        na.omit
    return(aba_summary_by_date)
}

#' Helper function that generates an ABA report summarized by queue and day of week, including total/overall
#' 
#' @param x A dataframe from CUIC, cleaned
#' @return A dataframe summarizing by queue and day of week, including total/overall
generate_aba_summary_by_weekday <- function(report_from_cuic_clean) {
    aba_summary_by_weekday_by_queue <-
        report_from_cuic_clean %>%
            group_by(precision_queue, weekday) %>%
            summarise(
                service_level_mean = mean(na.omit(service_level)),
                calls_offered_total = sum(na.omit(calls_offered)),
                calls_answered_total = sum(na.omit(calls_answered)),
                calls_abandoned_total = sum(na.omit(calls_abandoned)),
                handle_time_mean = mean(na.omit(handle_time_mean)),
                speed_of_answer_mean = mean(na.omit(speed_of_answer_mean)),
                hold_time_mean = mean(na.omit(hold_time))
            ) %>%
            ungroup() %>%
            mutate(aba = calls_abandoned_total/calls_offered_total)
    aba_summary_by_weekday_overall <-
        report_from_cuic_clean %>%
            group_by(weekday) %>%
            summarise(
                service_level_mean = mean(na.omit(service_level)),
                calls_offered_total = sum(na.omit(calls_offered)),
                calls_answered_total = sum(na.omit(calls_answered)),
                calls_abandoned_total = sum(na.omit(calls_abandoned)),
                handle_time_mean = mean(na.omit(handle_time_mean)),
                speed_of_answer_mean = mean(na.omit(speed_of_answer_mean)),
                hold_time_mean = mean(na.omit(hold_time))
            ) %>%
            ungroup() %>%
            mutate(aba = calls_abandoned_total/calls_offered_total,
                    precision_queue = "total_overall")
    aba_summary_by_weekday_by_queue <- 
        aba_summary_by_weekday_by_queue[c(1,2,10,3,4,5,6,7,8,9)]
    aba_summary_by_weekday_overall <- 
        aba_summary_by_weekday_overall[c(10,1,2,9,3,4,5,6,7,8)]
    aba_summary_by_weekday <-
        rbind(aba_summary_by_weekday_by_queue,
                aba_summary_by_weekday_overall) %>%
                na.omit
    return(aba_summary_by_weekday)
}

#' Helper function that generates an ABA report summarized by queue and hour of day, including total/overall
#' 
#' @param x A dataframe from CUIC, cleaned
#' @return A dataframe summarizing by queue and hour of day, including total/overall
generate_aba_summary_by_hour <- function(report_from_cuic_clean) {
    aba_summary_by_hour_by_queue <-
        report_from_cuic_clean %>%
            group_by(precision_queue, hour) %>%
            summarise(
                service_level_mean = mean(na.omit(service_level)),
                calls_offered_total = sum(na.omit(calls_offered)),
                calls_answered_total = sum(na.omit(calls_answered)),
                calls_abandoned_total = sum(na.omit(calls_abandoned)),
                handle_time_mean = mean(na.omit(handle_time_mean)),
                speed_of_answer_mean = mean(na.omit(speed_of_answer_mean)),
                hold_time_mean = mean(na.omit(hold_time))
            ) %>%
            ungroup() %>%
            mutate(aba = calls_abandoned_total/calls_offered_total)
    aba_summary_by_hour_overall <-
        report_from_cuic_clean %>%
            group_by(hour) %>%
            summarise(
                service_level_mean = mean(na.omit(service_level)),
                calls_offered_total = mean(na.omit(calls_offered)),
                calls_answered_total = mean(na.omit(calls_answered)),
                calls_abandoned_total = mean(na.omit(calls_abandoned)),
                handle_time_mean = mean(na.omit(handle_time_mean)),
                speed_of_answer_mean = mean(na.omit(speed_of_answer_mean)),
                hold_time_mean = mean(na.omit(hold_time))
            ) %>%
            ungroup() %>%
            mutate(aba = calls_abandoned_total/calls_offered_total,
                    precision_queue = "total_overall")
    aba_summary_by_hour_by_queue <- 
        aba_summary_by_hour_by_queue[c(1,2,10,3,4,5,6,7,8,9)]
    aba_summary_by_hour_overall <- 
        aba_summary_by_hour_overall[c(10,1,2,9,3,4,5,6,7,8)]
    aba_summary_by_hour <-
        rbind(aba_summary_by_hour_by_queue,
                aba_summary_by_hour_overall)
    return(aba_summary_by_hour)
}

#' Helper function that generates an ABA report summarized by queue, day of week, and hour of day, including total/overall
#' 
#' @param x A dataframe from CUIC, cleaned
#' @return A dataframe summarizing by queue, day of week, and hour of day, including total/overall
generate_aba_summary_by_date_by_hour <- function(report_from_cuic_clean) {
    aba_summary_by_date_by_hour_by_queue <-
        report_from_cuic_clean %>%
            group_by(precision_queue, date, hour) %>%
            summarise(
                service_level_mean = mean(na.omit(service_level)),
                calls_offered_total = sum(na.omit(calls_offered)),
                calls_answered_total = sum(na.omit(calls_answered)),
                calls_abandoned_total = sum(na.omit(calls_abandoned)),
                handle_time_mean = mean(na.omit(handle_time_mean)),
                speed_of_answer_mean = mean(na.omit(speed_of_answer_mean)),
                hold_time_mean = mean(na.omit(hold_time))
            ) %>%
            ungroup() %>%
            mutate(weekday = weekdays(date),
                aba = calls_abandoned_total/calls_offered_total)
    aba_summary_by_date_by_hour_overall <-
        report_from_cuic_clean %>%
            group_by(date, hour) %>%
            summarise(
                service_level_mean = mean(na.omit(service_level)),
                calls_offered_total = sum(na.omit(calls_offered)),
                calls_answered_total = sum(na.omit(calls_answered)),
                calls_abandoned_total = sum(na.omit(calls_abandoned)),
                handle_time_mean = mean(na.omit(handle_time_mean)),
                speed_of_answer_mean = mean(na.omit(speed_of_answer_mean)),
                hold_time_mean = mean(na.omit(hold_time))
            ) %>%
            ungroup() %>%
            mutate(weekday = weekdays(date),
                    aba = calls_abandoned_total/calls_offered_total,
                    precision_queue = "total_overall")
    aba_summary_by_date_by_hour_by_queue <- 
        aba_summary_by_date_by_hour_by_queue[c(1,2,3,11,12,4,5,6,7,8,9,10)]
    aba_summary_by_date_by_hour_overall <- 
        aba_summary_by_date_by_hour_overall[c(12,1,2,3,10,11,4,5,6,7,8,9)]
    aba_summary_by_date_by_hour <-
        rbind(aba_summary_by_date_by_hour_by_queue,
                aba_summary_by_date_by_hour_overall)
    return(aba_summary_by_date_by_hour)
}

generate_aba_summary_by_15_min <- function(report_from_cuic_clean) {
    aba_summary_by_15_min_by_queue <-
        report_from_cuic_clean %>%
            group_by(precision_queue, time) %>%
            summarise(
                service_level_mean = mean(service_level, na.rm = TRUE),
                calls_offered_median = median(calls_offered, na.rm = TRUE),
                calls_answered_median = median(calls_answered, na.rm = TRUE),
                calls_abandoned_median = median(calls_abandoned, na.rm = TRUE),
                handle_time_mean = mean(handle_time_mean, na.rm = TRUE),
                speed_of_answer_mean = mean(speed_of_answer_mean, na.rm = TRUE),
                hold_time_mean = mean(hold_time, na.rm = TRUE)
            ) %>%
            ungroup() %>%
            mutate(aba_median = calls_abandoned_median/calls_offered_median)
    aba_summary_by_15_min_overall <-
        report_from_cuic_clean %>%
            group_by(time) %>%
            summarise(
                service_level_mean = mean(service_level, na.rm = TRUE),
                calls_offered_median = median(calls_offered, na.rm = TRUE),
                calls_answered_median = median(calls_answered, na.rm = TRUE),
                calls_abandoned_median = median(calls_abandoned, na.rm = TRUE),
                handle_time_mean = mean(handle_time_mean, na.rm = TRUE),
                speed_of_answer_mean = mean(speed_of_answer_mean, na.rm = TRUE),
                hold_time_mean = mean(hold_time, na.rm = TRUE)
            ) %>%
            ungroup() %>%
            mutate(aba_median = calls_abandoned_median/calls_offered_median,
                    precision_queue = "total_overall")
    aba_summary_by_15_min_by_queue <- 
        aba_summary_by_15_min_by_queue[c(1,2,10,3,4,5,6,7,8,9)]
    aba_summary_by_15_min_overall <- 
        aba_summary_by_15_min_overall[c(10,1,2,9,3,4,5,6,7,8)]
    aba_summary_by_15_min <-
        rbind(aba_summary_by_15_min_by_queue,
                aba_summary_by_15_min_overall)
    return(aba_summary_by_15_min)
}

generate_aba_summary_by_date_by_15_min <- function(report_from_cuic_clean) {
    aba_summary_by_date_by_15_min_by_queue <-
        report_from_cuic_clean %>%
            group_by(precision_queue, date_time) %>%
            summarise(
                service_level_mean = mean(service_level, na.rm = TRUE),
                calls_offered_total = sum(calls_offered, na.rm = TRUE),
                calls_answered_total = sum(calls_answered, na.rm = TRUE),
                calls_abandoned_total = sum(calls_abandoned, na.rm = TRUE),
                handle_time_mean = mean(handle_time_mean, na.rm = TRUE),
                speed_of_answer_mean = mean(speed_of_answer_mean, na.rm = TRUE),
                hold_time_mean = mean(hold_time, na.rm = TRUE)
            ) %>%
            ungroup() %>%
            mutate(aba = calls_abandoned_total/calls_offered_total)
    aba_summary_by_date_by_15_min_overall <-
        report_from_cuic_clean %>%
            group_by(date_time) %>%
            summarise(
                service_level_mean = mean(service_level, na.rm = TRUE),
                calls_offered_total = sum(calls_offered, na.rm = TRUE),
                calls_answered_total = sum(calls_answered, na.rm = TRUE),
                calls_abandoned_total = sum(calls_abandoned, na.rm = TRUE),
                handle_time_mean = mean(handle_time_mean, na.rm = TRUE),
                speed_of_answer_mean = mean(speed_of_answer_mean, na.rm = TRUE),
                hold_time_mean = mean(hold_time, na.rm = TRUE)
            ) %>%
            ungroup() %>%
            mutate(aba = calls_abandoned_total/calls_offered_total,
                    precision_queue = "total_overall")
    aba_summary_by_date_by_15_min_by_queue <- 
        aba_summary_by_date_by_15_min_by_queue[c(1,2,10,3,4,5,6,7,8,9)]
    aba_summary_by_date_by_15_min_overall <- 
        aba_summary_by_date_by_15_min_overall[c(10,1,2,9,3,4,5,6,7,8)]
    aba_summary_by_date_by_15_min <-
        rbind(aba_summary_by_date_by_15_min_by_queue,
                aba_summary_by_date_by_15_min_overall)
    return(aba_summary_by_date_by_15_min)
}

prettify_aba_standard_summary <- function(aba_standard_summary) {
    aba_standard_summary_pretty <- aba_standard_summary %>%
        mutate(
            across(
                c(handle_time_mean, speed_of_answer_mean, hold_time_mean), 
                function(x) {x %>% format("%H:%M:%S") %>% hms}
                ),
            across(
                c(service_level_mean, aba), 
                function(x) {scales::percent(x, accuracy = 0.1)}
                ),
            across(
                c(calls_offered_total, calls_answered_total, calls_abandoned_total), 
                function(x) {prettyNum_tp(x)}
                )
            )
    return(aba_standard_summary_pretty)
}

#' Generates an ABA report summarized by queue and a set of options by choice
#' 
#' @param x A dataframe from CUIC, uncleaned
#' @param type The type of report to generate, options are "standard" (default), "by_date", "by_hour", "by_date_and_hour", "by_15_min", "by_date_and_15_min", "by_weekday"
#' @param queue The queue to generate for (CLSMA,INTMED,GIM,CIM,STUDENT,TRAVEL,CORP,PED,DERM). Use "all" to show all queues and the overall numbers (the default) and "overall" to only show the overall numbers.
#' @return A dataframe
#' @example generate_aba_summary(aba_report_2021-07-01)
#' @example generate_aba_summary(aba_report_2021-07-01, type = "by_date")
#' @example generate_aba_summary(aba_report_2021-07-01, type = "by_date_and_hour", queue = "CLSMA")
generate_aba_summary <- function(
    data, 
    type = "standard", 
    queue = "all", 
    new_date_type = FALSE,
    exclude_saturday = FALSE, 
    presentation_view = FALSE
    ) {
    cleaned_report <- clean_aba_report_from_cuic(data, new_date_type = new_date_type)
    if (exclude_saturday == TRUE) {
        cleaned_report <- cleaned_report %>% filter(weekday != "Saturday")
    } else {}
    if (type == "standard" & presentation_view == FALSE) {
        aba_summary <- 
            generate_aba_summary_standard(cleaned_report)
    } else if (type == "standard" & presentation_view == TRUE) {
        aba_summary <- 
            generate_aba_summary_standard(cleaned_report) %>%
                prettify_aba_standard_summary()
    } else if (type == "by_date") {
        aba_summary <-
            generate_aba_summary_by_date(cleaned_report)
    } else if (type == "by_hour") {
        aba_summary <-
            generate_aba_summary_by_hour(cleaned_report)
    } else if (type == "by_date_and_hour") {
        aba_summary <-
            generate_aba_summary_by_date_by_hour(cleaned_report)
    } else if (type == "by_weekday") {
        aba_summary <-
            generate_aba_summary_by_weekday(cleaned_report)
    } else if (type == "by_15_min") {
        aba_summary <-
            generate_aba_summary_by_15_min(cleaned_report)
    } else if (type == "by_date_and_15_min") {
        aba_summary <-
            generate_aba_summary_by_date_by_15_min(cleaned_report)
    }
    if (queue == "all") {
        summary_filtered_by_queue <-
            aba_summary
    } else {
        summary_filtered_by_queue <-
            aba_summary %>%
                filter(grepl(collapse_vector(queue), precision_queue))
    }
    return(summary_filtered_by_queue)
}

#' Generates and saves an Excel workbook containing all five types of ABA reports to the current working directory
#' 
#' @param x A dataframe from CUIC, uncleaned
#' @param report_date_or_month The date of the report, or the month of the report (format Jun 21), defaults to the current date
#' @param report_name Filename suffix, defaults to "export"
#' @return Workbook saved and success message printed
generate_and_save_aba_summary_combined <- function(report_from_cuic,
                                                   new_date_type = FALSE,
                                                   report_date_or_month = Sys.Date(), 
                                                   report_name = "export") {
    summary_standard <-
        generate_aba_summary(
            report_from_cuic,
            new_date_type = new_date_type,
            "standard",
            "all"
        )
    summary_by_date <- 
        generate_aba_summary(
            report_from_cuic,
            new_date_type = new_date_type,
            "by_date",
            "all"
        ) %>% na.omit
    summary_by_weekday <-
        generate_aba_summary(
            report_from_cuic,
            new_date_type = new_date_type,
            "by_weekday",
            "all"
        ) %>% na.omit
    summary_by_hour <-
        generate_aba_summary(
            report_from_cuic,
            new_date_type = new_date_type,
            "by_hour",
            "all"
        ) %>% na.omit
    summary_by_date_and_hour <-
        generate_aba_summary(
            report_from_cuic,
            new_date_type = new_date_type,
            "by_date_and_hour",
            "all"
        ) %>% na.omit
    summary_by_15_min <-
        generate_aba_summary(
            report_from_cuic,
            new_date_type = new_date_type,
            "by_15_min",
            "all"
        ) %>% na.omit
    summary_by_date_and_15_min <-
        generate_aba_summary(
            report_from_cuic,
            new_date_type = new_date_type,
            "by_date_and_15_min",
            "all"
        ) %>% na.omit
    report_sheet_list <- 
        lst(
            summary_standard,
            summary_by_date,
            summary_by_weekday,
            summary_by_hour,
            summary_by_date_and_hour,
            summary_by_15_min,
            summary_by_date_and_15_min
        )
    write.xlsx(
        report_sheet_list,
        paste0(
            "aba_combined_report",
            "_",
            report_name,
            ".xlsx"
        ),
        overwrite = TRUE
    )
    return(paste0("Success! Data written to file."))
}

## functions to automatically generate interactive plots for dashboards

plot_aba_summary <- function(aba_data, 
                                type, 
                                queue_search = NA, 
                                exclude_queues = NA) {
    aba_data <- aba_data %>%
        pivot_longer(!c(Queue, Month), names_to = "variable")
    if (!is.na(queue_search[1])) {
        queue_search <- collapse_vector(queue_search)
        aba_data <- aba_data %>%
            filter(grepl(queue_search, Queue))
    } else {}
    if (!is.na(exclude_queues[1])) {
        exclude_queues <- collapse_vector(exclude_queues)
        aba_data <- aba_data %>%
            filter(!grepl(exclude_queues, Queue))
    } else {}
    if (type == "calls") {
        plot <- aba_data %>%
            filter(grepl("Calls", variable)) %>%
            ggplot(aes(x = Month, y = value)) +
                geom_path(aes(color = Queue)) +
                scale_y_continuous(
                    # breaks = function(x) {seq(0, max(x), by = 2500)},
                    labels = function(x) {round(x, digits = 0)}
                    ) +
                facet_wrap(. ~ variable, scales = "free_y")
    } else if (type == "service_level") {
        plot <- aba_data %>%
            filter(variable == "Service Level (Avg)") %>%
            ggplot(aes(x = Month, y = value)) +
                geom_path(aes(color = Queue)) +
                scale_y_continuous(labels = function(x) {scales::percent(x, accuracy = 1)})
    } else if (type == "times") {
        plot <- aba_data %>%
            filter(variable == "Answer Speed" | variable == "Handle Time") %>%
            ggplot(aes(x = Month, y = value)) +
                geom_line(aes(color = Queue)) +
                scale_y_continuous(
                    labels = function(x) {x %>% 
                        seconds_to_period %>% 
                        as_datetime %>% 
                        format("%M:%S")},
                    breaks = function(x) {seq(0, max(x+60), by = 60)}
                    ) +
                facet_wrap(. ~ variable, scales = "free_y")
    } else if (type == "aba") {
        plot <- aba_data %>%
            filter(variable == "ABA") %>%
            ggplot(aes(x = Month, y = value)) +
                geom_line(aes(color = Queue)) +
                scale_y_continuous(
                    labels = function(x) {scales::percent(x, accuracy = 1)},
                    breaks = function(x) {seq(0, max(x), by = 0.05)}
                    )
    } else if (type == "service_level_aba") {
        plot <- aba_data %>%
            filter(variable == "ABA" | variable == "Service Level (Avg)") %>%
            ggplot(aes(x = Month, y = value)) +
                geom_line(aes(color = Queue)) +
                facet_wrap(variable ~ ., scales = "free_y") +
                scale_y_continuous(
                    labels = function(x) {scales::percent(x)}
                    )
    }
    plot <- plot + 
        theme_bw() + 
        expand_limits(y = 0) +
        scale_x_date(date_breaks = "1 month", date_labels = "%b %y") +
        geom_smooth(method = "lm", se = F, linetype = "dotted", size = 0.25)
    return(plot)
}

customize_hovertext <- function(plotly_object, plot_type) {
  chart_font <- list(family = "arial", size = 12)
  # plotly_object$x$layout$showlegend <- FALSE
  get_hovertext_for_chart_number <- function(chart_number, type = plot_type) {
    if (class(plotly_object$x$data[[1]]$x) != "Date") {
        hovertext_x <- 
      format(convertToDate(plotly_object$x$data[[chart_number]]$x, origin = "1970-01-01"), "%b %Y")
    } else {
        hovertext_x <- 
      format(plotly_object$x$data[[chart_number]]$x, "%b %Y")
    }
    if (type == "aba") {
      hovertext_y <- plotly_object$x$data[[chart_number]]$y %>% as.numeric %>% scales::percent(accuracy = 1)
    #   hovertext_perc <- "%"
    } else if (type == "calls") {
      hovertext_y <- round(plotly_object$x$data[[chart_number]]$y, digits = 0)
    #   hovertext_perc <- NA
    } else if (type == "times") {
      hovertext_y <- 
        round(seconds_to_period(plotly_object$x$data[[chart_number]]$y), digits = 0)
    #   hovertext_perc <- NA
    }
    hovertext <- paste0(hovertext_x, ": ", hovertext_y)
    return(hovertext)
  }
  x_axis_format <- list(tickfont = chart_font)
  if (grepl("aba", plot_type)[1] == TRUE) {
    for (i in c(1,3)) {
      plotly_object$x$data[[i]]$text <- get_hovertext_for_chart_number(i)
    }
    plotly_object$x$data[[3]]$showlegend <- TRUE
    plotly_object$x$data[[3]]$name <- "Linear Trend"
    plotly_object$x$layout$annotations[[1]]$font$family <- "arial"
    plotly_object <- plotly_object %>% 
        layout(
            xaxis = x_axis_format,
            yaxis = list(
                tickfont = chart_font, 
                rangemode = "tozero", 
                tickformat = "%"
                    )
                )
  } else if (grepl("calls", plot_type)[1] == TRUE) {
    for (i in c(1:3, 7:9)) {
      plotly_object$x$data[[i]]$text <- get_hovertext_for_chart_number(i)
    }
    for (i in 7:9) {
      names <- c(1,2,3,4,5,6,"Offered", "Abandoned", "Answered")
      plotly_object$x$data[[i]]$showlegend <- TRUE
      plotly_object$x$data[[i]]$name <- paste0("Linear Trend", " (", names[i], ")")
    }
    for (i in 1:3) {
      plotly_object$x$layout$annotations[[i]]$font$family <- "arial"
    }
    y_axis_format_calls <- list(tickfont = chart_font, rangemode = "tozero", tickformat = ",d")
    plotly_object <- plotly_object %>% layout(
      xaxis = x_axis_format,
      yaxis = y_axis_format_calls,
      xaxis2 = x_axis_format,
      yaxis2 = y_axis_format_calls,
      xaxis3 = x_axis_format,
      yaxis3 = y_axis_format_calls
    )
  } else if (grepl("times", plot_type)[1] == TRUE) {
    for (i in 1:6) {
      plotly_object$x$data[[i]]$text <- get_hovertext_for_chart_number(i)
    }
    for (i in 4:6) {
      names <- c(1,2,3,"", "Answer Speed", "Handle Time")
      plotly_object$x$data[[i]]$showlegend <- TRUE
      plotly_object$x$data[[i]]$name <- paste0("Linear Trend", " (", names[i], ")")
    }
    for (i in 1:2) {
      plotly_object$x$layout$annotations[[i]]$font$family <- "arial"
    }
    y_axis_format_times <- 
      list(
          tickfont = chart_font
           )
    plotly_object <- plotly_object %>% layout(
      xaxis = x_axis_format,
      yaxis = y_axis_format_times,
      xaxis2 = x_axis_format,
      yaxis2 = y_axis_format_times
    )
  }
  plotly_object <- plotly_object %>% 
    layout(hovermode = "x", 
           legend = list(font = chart_font))
  return(plotly_object)
}

plotly_aba_summary <- function(aba_data, 
                               type, 
                               queue_search = NA, 
                               exclude_queues = NA) {
  plot_aba <- plot_aba_summary(aba_data, type, queue_search, exclude_queues)
  plotly_aba <- customize_hovertext(ggplotly(plot_aba, dynamicTicks = FALSE), type)
  return(plotly_aba)
}