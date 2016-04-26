package com.file.dao;

import java.util.List;

import javax.annotation.Resource;
import javax.sql.DataSource;

import org.springframework.jdbc.core.BeanPropertyRowMapper;
import org.springframework.jdbc.core.namedparam.SqlParameterSource;
import org.springframework.jdbc.core.namedparam.SqlParameterSourceUtils;
import org.springframework.jdbc.core.simple.SimpleJdbcTemplate;
import org.springframework.stereotype.Repository;

import com.file.forms.Computer;
@Repository("poiDao")
public class PoiDao {

	private SimpleJdbcTemplate jdbcTemplate;  
    @Resource(name = "dataSource")  
    public void setDataSource(DataSource dataSource) {  
        this.jdbcTemplate = new SimpleJdbcTemplate(dataSource);  
    }  
  
    public List<Computer> getComputer() {  
        String sql = "SELECT * FROM Computer";  
        return jdbcTemplate.query(sql, new BeanPropertyRowMapper<Computer>(  
                Computer.class));  
  
    }  
  
    public int[] insertComputer(List<Computer> list) {  
        String sql = "INSERT INTO Computer (BRAND,CPU,GPU,MEMORY,PRICE)VALUES(:brand,:cpu,:gpu,:memory,:price)";  
        SqlParameterSource[] batch = SqlParameterSourceUtils.createBatch(list  
                .toArray());  
        return jdbcTemplate.batchUpdate(sql, batch);  
  
    }  
}
