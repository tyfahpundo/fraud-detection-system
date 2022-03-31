package com.bezkoder.spring.files.excel.model;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;

@Entity
@Table(name = "finance")
public class Finance {

  @Id
  @Column(name = "id")
  private long id;
  @Column(name = "description")
  private String description;

  @Column(name = "sales")
  private Long sales;

  public Finance() {
  }

  public Finance(long id, String description, Long sales) {
    this.id = id;
    this.description = description;
    this.sales = sales;
  }

  public long getId() {
    return id;
  }

  public void setId(long id) {
    this.id = id;
  }

  public String getDescription() {
    return description;
  }

  public void setDescription(String description) {
    this.description = description;
  }

  public Long getSales() {
    return sales;
  }

  public void setSales(Long sales) {
    this.sales = sales;
  }
}
