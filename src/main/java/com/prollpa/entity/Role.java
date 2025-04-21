package com.prollpa.entity;

import java.util.Set;

import com.fasterxml.jackson.annotation.JsonIgnore;

import jakarta.persistence.CascadeType;
import jakarta.persistence.Column;
import jakarta.persistence.Entity;

import jakarta.persistence.GeneratedValue;
import jakarta.persistence.GenerationType;
import jakarta.persistence.Id;
import jakarta.persistence.OneToMany;
import jakarta.persistence.Table;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@AllArgsConstructor
@NoArgsConstructor
@Entity
@Table(name ="t_role")
public class Role  {
	@Id
	@GeneratedValue(strategy = GenerationType.IDENTITY)
	@Column(name = "role_id")
	private long roleId;
	private String role;
	@OneToMany(mappedBy = "role", cascade = CascadeType.ALL)
	@JsonIgnore
    private Set<UserVscRole> userVscRoles;
	
	@OneToMany(mappedBy = "role", cascade = CascadeType.ALL)
	@JsonIgnore
    private Set<UserRole> userRole;
	

}
