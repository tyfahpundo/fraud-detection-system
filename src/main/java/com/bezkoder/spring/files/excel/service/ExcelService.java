package com.bezkoder.spring.files.excel.service;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import com.bezkoder.spring.files.excel.helper.ExcelHelper;
import com.bezkoder.spring.files.excel.model.Finance;
import com.bezkoder.spring.files.excel.repository.FinanceRepository;

@Service
public class ExcelService {
  @Autowired
  FinanceRepository repository;

  public void save(MultipartFile file) {
    try {
      List<Finance> finances = ExcelHelper.excelToTutorials(file.getInputStream());
      repository.saveAll(finances);
    } catch (IOException e) {
      throw new RuntimeException("fail to store excel data: " + e.getMessage());
    }
  }

  public ByteArrayInputStream load() {
    List<Finance> finances = repository.findAll();

    ByteArrayInputStream in = ExcelHelper.tutorialsToExcel(finances);
    return in;
  }

  public List<Finance> getAllStatements() {
    return repository.findAll();
  }
}
