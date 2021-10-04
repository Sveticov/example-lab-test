import { TestBed } from '@angular/core/testing';

import { SecondServiceService } from './second-service.service';

describe('SecondServiceService', () => {
  let service: SecondServiceService;

  beforeEach(() => {
    TestBed.configureTestingModule({});
    service = TestBed.inject(SecondServiceService);
  });

  it('should be created', () => {
    expect(service).toBeTruthy();
  });
});
