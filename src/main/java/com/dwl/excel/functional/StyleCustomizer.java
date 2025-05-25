package com.dwl.excel.functional;

@FunctionalInterface
public interface StyleCustomizer<T> {
  void customize(T t);
}
