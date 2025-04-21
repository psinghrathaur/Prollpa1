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
import jakarta.validation.constraints.NotNull;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
@Entity
@Data
@Table(name ="t_user_mas")
public class User   {
	 @Id
	 @GeneratedValue(strategy = GenerationType.IDENTITY)  // Auto-generate userId
	 @Column(name="user_id")
	 private long userId;
	 @NotNull(message = "Username cannot be null")  // Enforce non-null username
	 private String username;
     @NotNull(message = "Email cannot be null")  // Enforce non-null email
	 private String email;
     @NotNull(message = "Name cannot be null")  // Enforce non-null name
     private String name;
     @OneToMany(mappedBy = "user", cascade = CascadeType.ALL)
     @JsonIgnore
     private Set<UserVscRole> userVscRoles;
     @OneToMany(mappedBy = "user", cascade = CascadeType.ALL)
     @JsonIgnore
     private Set<UserRole> userRole;
}
