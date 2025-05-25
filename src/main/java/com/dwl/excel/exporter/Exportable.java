package com.dwl.excel.exporter;

/**
 * Classes implementing this interface are considered eligible for data export, such as to Excel.
 *
 * Custom objects that need to support export functionality must implement this interface.
 * The export module uses the presence of this interface to determine whether a class is exportable.
 */
public interface Exportable {}
