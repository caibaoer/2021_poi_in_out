package com.hp.entity;

import lombok.Data;

import java.lang.reflect.Field;
import java.util.Objects;

@Data
public class ExcelEntityField
{
	private String columnName;
	
	private int columnIndex;
	
	private int columnType;
	
	private Field field;

	private boolean nullable;
	//最大长度
	private int maxLength;
	//最小长度
	private int minLength;


	@Override
	public boolean equals(Object o) {
		if (this == o) return true;
		if (o == null || getClass() != o.getClass()) return false;
		ExcelEntityField that = (ExcelEntityField) o;
		return columnIndex == that.columnIndex &&
				columnType == that.columnType &&
				Objects.equals(columnName, that.columnName);
	}

	@Override
	public int hashCode() {
		return Objects.hash(columnName, columnIndex, columnType, field);
	}
}
