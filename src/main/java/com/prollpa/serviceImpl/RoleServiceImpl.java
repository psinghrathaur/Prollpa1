package com.prollpa.serviceImpl;

import java.util.List;

import org.modelmapper.ModelMapper;
import org.springframework.stereotype.Service;

import com.prollpa.entity.Role;
import com.prollpa.exception.ResourceNotFoundException;
import com.prollpa.payload.RoleDto;
import com.prollpa.repository.RoleRepository;
import com.prollpa.service.RoleService;

@Service
public class RoleServiceImpl implements RoleService{
	private RoleRepository roleRepository;
	private final ModelMapper modelMapper;

	public RoleServiceImpl(RoleRepository roleRepository,ModelMapper modelMapper) {
		this.modelMapper = modelMapper;
		this.roleRepository = roleRepository;
	}
	
	public Role saveRole(Role role){
		return roleRepository.save(role);
        
	}
	public List<Role> getRoleList(){
		return roleRepository.findAll();
	}

	@Override
	public Role getRoleByRoleId(long roleId) {
		
		return roleRepository.findById(roleId).orElseThrow(
			()-> new ResourceNotFoundException("role not found with this id "+ roleId));
	}
	public  RoleDto mapRoletoDto(Role role) {
			return modelMapper.map(role,RoleDto.class);
			
		}

	@Override
	public Role getRoleByRoleId(Long roleId) {
		return roleRepository.findById(roleId).orElseThrow(
				()-> new ResourceNotFoundException("roleNot found with this role Id"));
		
	}
	}
	
	
	

