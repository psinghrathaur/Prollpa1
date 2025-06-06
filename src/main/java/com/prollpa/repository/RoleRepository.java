package com.prollpa.repository;

import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import com.prollpa.entity.Role;

@Repository
public interface RoleRepository extends JpaRepository<Role, Long>{
   Optional<Role> findByRole(String role);
}
