package javatime;

import java.time.YearMonth;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.collections4.CollectionUtils;

public class TestYearMonth {
	public static void main(String[] args) {
        String shipYearFrom = "2019";
        String shipMonthFrom = "01";
        String shipYearTo = "2019";
        String shipMonthTo = "02";
		YearMonth start = YearMonth.of(Integer.parseInt(shipYearFrom), Integer.parseInt(shipMonthFrom));
		YearMonth end = YearMonth.of(Integer.parseInt(shipYearTo), Integer.parseInt(shipMonthTo));
		List<String> yearMonths = new ArrayList<String>();
		yearMonths.add("'"+start.format(DateTimeFormatter.ofPattern("yyyy-MM"))+"'");
		while (start.isBefore(end)) {
			start = start.plusMonths(1);
			yearMonths.add("'"+start.format(DateTimeFormatter.ofPattern("yyyy-MM"))+"'");
		}
		System.out.println(yearMonths);
		String yearMonthsArray = "";
		if (CollectionUtils.isNotEmpty(yearMonths)) {
			yearMonthsArray = "(" + String.join(",", yearMonths) + ") ";
		}
		System.out.println("yearMonthsArray : " + yearMonthsArray);
	}
}
