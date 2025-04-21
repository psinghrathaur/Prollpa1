package com.prollpa.serviceImpl;

import java.util.List;

import org.springframework.stereotype.Service;

import com.prollpa.entity.VSC;
import com.prollpa.exception.ResourceNotFoundException;
import com.prollpa.repository.VSCRepository;
import com.prollpa.service.VSCService;
@Service
public class VSCServiceImpl implements VSCService{
	
	private VSCRepository vscRepository;
	public VSCServiceImpl(VSCRepository vscRepository) {
		this.vscRepository=vscRepository;
	}

	@Override
	public VSC saveVSC(VSC vsc) {
		// TODO Auto-generated method stub
		VSC SaveVSC = vscRepository.save(vsc);
		return SaveVSC;
	}

	@Override
	public VSC getVSCById(Long vscId) {
		// TODO Auto-generated method stub
		
		return vscRepository.findById(vscId).orElseThrow(
				()-> new ResourceNotFoundException("VSC not found with this id "+vscId));
		
	}

	@Override
	public List<VSC> getVSCList() {
		return vscRepository.findAll();
	}

}
