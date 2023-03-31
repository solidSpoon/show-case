package xyz.solidspoon.showCase;

import java.lang.annotation.*;

@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface TableColumnIndex {
	String value();
}
