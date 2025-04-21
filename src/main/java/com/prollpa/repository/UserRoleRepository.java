package com.prollpa.repository;

import java.util.List;

import org.springframework.data.jpa.repository.JpaRepository;

import com.prollpa.entity.UserRole;
import com.prollpa.entity.UserVscRole;

public interface UserRoleRepository extends JpaRepository<UserRole, Long>{

	List<UserRole> findByUser_userId(Long userId);
	
}
