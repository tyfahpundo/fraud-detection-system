package com.bezkoder.spring.files.excel.repository;

import org.springframework.data.jpa.repository.JpaRepository;

import com.bezkoder.spring.files.excel.model.Finance;

public interface FinanceRepository extends JpaRepository<Finance, Long> {
}
