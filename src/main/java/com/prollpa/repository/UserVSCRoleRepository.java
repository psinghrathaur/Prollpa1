package com.prollpa.repository;

import java.util.List;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.jpa.repository.config.EnableJpaRepositories;

import com.prollpa.entity.UserVscRole;

public interface UserVSCRoleRepository extends JpaRepository<UserVscRole, Long>{
    List<UserVscRole> findByUser_userId(long userId);
    List<UserVscRole> findByRole_roleId(long roleId);
    List<UserVscRole> findByVsc_vscId(long vscId);
    
    
	 

}
