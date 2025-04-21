package com.prollpa.entity;

import jakarta.persistence.Entity;
import jakarta.persistence.GeneratedValue;
import jakarta.persistence.GenerationType;
import jakarta.persistence.Id;
import jakarta.persistence.JoinColumn;
import jakarta.persistence.ManyToOne;
import jakarta.persistence.Table;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Data
@Entity
@Table(name = "t_user_vsc_role")
public class UserVscRole  {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long userVscRoleId;

    // Many-to-One relationship with VSC
    @ManyToOne
    @JoinColumn(name = "vsc_id", referencedColumnName = "vsc_id", nullable = false)
    private VSC vsc;

    // Many-to-One relationship with Role
    @ManyToOne
    @JoinColumn(name = "role_id", referencedColumnName = "role_id", nullable = false)
    private Role role;

    // Many-to-One relationship with User
    @ManyToOne
    @JoinColumn(name = "user_id", referencedColumnName = "user_id", nullable = false)
    private User user;
}
