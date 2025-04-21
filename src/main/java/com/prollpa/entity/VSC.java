package com.prollpa.entity;

import java.util.List;
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
import lombok.Data;

@Entity
@Data
@Table(name ="t_vsc_code")
public class VSC  {
	@Id
	@GeneratedValue(strategy = GenerationType.IDENTITY)  // Auto-generate userId
	@Column(name = "vsc_id")
	private Long vscId;
	private String vscCenterName;
	@OneToMany(mappedBy = "vsc", cascade = CascadeType.ALL)
	@JsonIgnore
    private Set<UserVscRole> userVscRoles;

}
