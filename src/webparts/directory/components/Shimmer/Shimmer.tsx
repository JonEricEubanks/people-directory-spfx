import { makeStyles, Skeleton, SkeletonItem } from '@fluentui/react-components';
import React from 'react';
import { SKELETON_COUNT } from '../../../../constants';

const useSkeletonStyles = makeStyles({
  grid: {
    display: 'flex',
    flexWrap: 'wrap',
    justifyContent: 'center',
    gap: '16px',
    padding: '8px 0',
    marginTop: '10px',
  },
  card: {
    width: '240px',
    height: '290px',
    borderRadius: '12px',
    border: '1px solid #e8e8e8',
    overflow: 'hidden',
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    backgroundColor: '#ffffff',
    boxShadow: '0 2px 8px rgba(0,0,0,0.07)',
    flexShrink: 0,
  },
  header: {
    width: '100%',
    height: '72px',
    flexShrink: 0,
    borderRadius: '0',
  },
  avatarRing: {
    marginTop: '-40px',
    display: 'flex',
    justifyContent: 'center',
    position: 'relative',
    zIndex: 1,
    backgroundColor: '#ffffff',
    borderRadius: '50%',
    padding: '3px',
  },
  body: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    padding: '8px 14px 14px',
    gap: '8px',
    flex: 1,
    width: '100%',
    boxSizing: 'border-box',
  },
  nameLine: {
    width: '140px',
    height: '16px',
    borderRadius: '4px',
  },
  titleLine: {
    width: '110px',
    height: '10px',
    borderRadius: '4px',
  },
  deptPill: {
    width: '88px',
    height: '20px',
    borderRadius: '10px',
  },
  divider: {
    width: '32px',
    height: '2px',
    borderRadius: '2px',
    marginTop: '4px',
  },
  contactLine1: {
    width: '160px',
    height: '12px',
    borderRadius: '4px',
    alignSelf: 'flex-start',
  },
  contactLine2: {
    width: '130px',
    height: '12px',
    borderRadius: '4px',
    alignSelf: 'flex-start',
  },
});

const SkeletonCard = (): JSX.Element => {
  const sk = useSkeletonStyles();
  return (
    <Skeleton animation="pulse" aria-label="Loading person card">
      <div className={sk.card}>
        <SkeletonItem shape="rectangle" className={sk.header} />
        <div className={sk.avatarRing}>
          <SkeletonItem shape="circle" size={72} />
        </div>
        <div className={sk.body}>
          <SkeletonItem shape="rectangle" className={sk.nameLine} />
          <SkeletonItem shape="rectangle" className={sk.titleLine} />
          <SkeletonItem shape="rectangle" className={sk.deptPill} />
          <SkeletonItem shape="rectangle" className={sk.divider} />
          <SkeletonItem shape="rectangle" className={sk.contactLine1} />
          <SkeletonItem shape="rectangle" className={sk.contactLine2} />
        </div>
      </div>
    </Skeleton>
  );
};

export const Shimmer = (): JSX.Element => {
  const sk = useSkeletonStyles();
  return (
    <div className={sk.grid}>
      {Array(SKELETON_COUNT)
        .fill(0)
        .map((_, index) => (
          <SkeletonCard key={index} />
        ))}
    </div>
  );
};
