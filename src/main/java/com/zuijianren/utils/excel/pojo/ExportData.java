package com.zuijianren.utils.excel.pojo;

import java.util.List;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * @author: 姜辞旧
 * @since: 2022/3/18 9:41
 * @version: 1.0
 */
@Data
@AllArgsConstructor
public class ExportData<T> {

    private Class<T> clazz;
    private List<T> list;

}
